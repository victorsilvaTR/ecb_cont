VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Begin VB.Form frmManConceptoLibros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Conceptos por Libro"
   ClientHeight    =   5955
   ClientLeft      =   2610
   ClientTop       =   2940
   ClientWidth     =   7965
   Icon            =   "frmManConceptoLibros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7965
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5310
      Left            =   45
      TabIndex        =   6
      Top             =   420
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   9366
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consultas de Conceptos x Libros"
      TabPicture(0)   =   "frmManConceptoLibros.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Conceptos x Libors"
      TabPicture(1)   =   "frmManConceptoLibros.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4875
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   7665
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   3510
            Left            =   180
            TabIndex        =   2
            Top             =   1215
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   6191
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Libro"
            Columns(0).DataField=   "Lib_cDescripcion"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Diario"
            Columns(1).DataField=   "Lib_cTipoLibro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Codigo "
            Columns(2).DataField=   "Asl_cCodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Descripción del Concepto "
            Columns(3).DataField=   "Asl_cDescripcion"
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
            Columns(4).Caption=   "Por Defecto"
            Columns(4).DataField=   "Asl_cDefecto"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   16
            Columns(5)._MaxComboItems=   5
            Columns(5).ValueItems(0)._DefaultItem=   0
            Columns(5).ValueItems(0).Value=   "1"
            Columns(5).ValueItems(0).Value.vt=   8
            Columns(5).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(5).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(5).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(5).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(5).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(5).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(5).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(5).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(5).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(5).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(5).ValueItems(0).DisplayValue.vt=   9
            Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(5).ValueItems(1)._DefaultItem=   0
            Columns(5).ValueItems(1).Value=   "0"
            Columns(5).ValueItems(1).Value.vt=   8
            Columns(5).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(5).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(5).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(5).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(5).ValueItems(1).DisplayValue.vt=   9
            Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(5).ValueItems.Count=   2
            Columns(5).Caption=   "Detallado"
            Columns(5).DataField=   "Asl_cDetallado"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3360"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3281"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Merge=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1085"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1005"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8724"
            Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(17)=   "Column(2).Width=1640"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1561"
            Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(23)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(24)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(25)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(26)=   "Column(3).Width=5477"
            Splits(0)._ColumnProps(27)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(3)._WidthInPix=5398"
            Splits(0)._ColumnProps(29)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(32)=   "Column(4).Width=1482"
            Splits(0)._ColumnProps(33)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(4)._WidthInPix=1402"
            Splits(0)._ColumnProps(35)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=529"
            Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(38)=   "Column(5).Width=1402"
            Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=1323"
            Splits(0)._ColumnProps(41)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=529"
            Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
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
            HeadLines       =   2
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   2715
            TabIndex        =   0
            Top             =   510
            Width           =   3750
            _Version        =   65536
            _ExtentX        =   6615
            _ExtentY        =   556
            Caption         =   "frmManConceptoLibros.frx":0F02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManConceptoLibros.frx":0F6E
            Key             =   "frmManConceptoLibros.frx":0F8C
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
         Begin TDBText6Ctl.TDBText tdbtLibroBus 
            Height          =   315
            Left            =   2715
            TabIndex        =   1
            Top             =   840
            Width           =   3750
            _Version        =   65536
            _ExtentX        =   6615
            _ExtentY        =   556
            Caption         =   "frmManConceptoLibros.frx":0FDE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManConceptoLibros.frx":104A
            Key             =   "frmManConceptoLibros.frx":1068
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
            Index           =   0
            Left            =   480
            TabIndex        =   14
            Top             =   885
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción de Conepto"
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
            Left            =   480
            TabIndex        =   13
            Top             =   555
            Width           =   1995
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
            Left            =   480
            TabIndex        =   12
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   -74730
         TabIndex        =   7
         Top             =   540
         Width           =   6930
         Begin VB.CheckBox chkDefecto 
            Height          =   330
            Left            =   1770
            TabIndex        =   19
            Top             =   2925
            Width           =   240
         End
         Begin VB.CheckBox chkDetallado 
            Height          =   330
            Left            =   1770
            TabIndex        =   17
            Top             =   4590
            Visible         =   0   'False
            Width           =   240
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   1770
            TabIndex        =   4
            Top             =   1515
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   556
            Caption         =   "frmManConceptoLibros.frx":10BA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManConceptoLibros.frx":1126
            Key             =   "frmManConceptoLibros.frx":1144
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
            Left            =   1770
            TabIndex        =   5
            Tag             =   "_"
            Top             =   2055
            Width           =   4500
            _Version        =   65536
            _ExtentX        =   7937
            _ExtentY        =   556
            Caption         =   "frmManConceptoLibros.frx":1196
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManConceptoLibros.frx":1202
            Key             =   "frmManConceptoLibros.frx":1220
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
         Begin TrueOleDBList70.TDBCombo tdbcLibro 
            Height          =   300
            Left            =   1770
            TabIndex        =   3
            Tag             =   "enabled"
            Top             =   1095
            Width           =   3120
            _ExtentX        =   5503
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
            Splits(0)._ColumnProps(17)=   "Column(3).Width=2196"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2117"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2196"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2117"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
            _PropDict       =   $"frmManConceptoLibros.frx":1272
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
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
         Begin VB.Label Label3 
            Caption         =   $"frmManConceptoLibros.frx":12F9
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   600
            Index           =   1
            Left            =   2115
            TabIndex        =   22
            Top             =   2970
            Width           =   4560
         End
         Begin VB.Label Label3 
            Caption         =   "(Permite detallar los movimientos de caja asociados a este concepto, en el reporte del diario simplificado)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   600
            Index           =   0
            Left            =   2115
            TabIndex        =   21
            Top             =   4635
            Visible         =   0   'False
            Width           =   4560
         End
         Begin VB.Label Label2 
            Caption         =   "Concepto por Defecto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   2
            Left            =   360
            TabIndex        =   20
            Top             =   2970
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Mostrar Concepto Detallado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   4635
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Diario"
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
            Left            =   345
            TabIndex        =   15
            Top             =   1140
            Width           =   495
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
            Left            =   345
            TabIndex        =   10
            Top             =   2085
            Width           =   990
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
            Left            =   315
            TabIndex        =   9
            Top             =   1590
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
            Left            =   315
            TabIndex        =   8
            Top             =   795
            Width           =   1515
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   3615
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
            Picture         =   "frmManConceptoLibros.frx":138D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":1767
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":1B41
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":1F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":22F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":26CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":2AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":2E83
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":3E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":3FF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":4151
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":42AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":4405
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":455F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":46B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":4813
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":496D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":4F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":54A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":5A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":5FD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":656F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":6B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManConceptoLibros.frx":70A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   16
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir F7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar o Salir ESC"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManConceptoLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
'Dim lrstabla As Recordset
Dim lrsTabla As ADODB.Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim Control As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left + 15 - 300
            .Height = Me.Height - .Top - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 200
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 200
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 300
            .Height = Frame1.Height - .Top - 200
        End With
        
        '*** REDIMENSIONAR DETALLE
        Frame2.Width = SSTCentroCosto.Width - 500
        Frame2.Height = SSTCentroCosto.Height - 800
        
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
            ReDim lArrMnt(11) As Variant
            lArrMnt(0) = "ELIMINAR"     ' Accion
            lArrMnt(1) = gsEmpresa      ' Codigo de Empresa
            lArrMnt(2) = gsAnio         ' Codigo de Plantilla
            lArrMnt(3) = tdbgCostos.Columns(1).Value    ' Tipo de Libro
            lArrMnt(4) = tdbgCostos.Columns(2).Value    ' Codigo Operacion
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaConceptosXLibros", lArrMnt(), True) = False Then
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
    
    tdbcLibro.Enabled = False
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
    Else
        lRegElim = False
    End If
    
    tdbcLibro.Enabled = False
End Sub

Private Sub ManNuevo()
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    
    chkDefecto.Value = vbUnchecked
    chkDetallado.Value = vbUnchecked
    
    tdbcLibro.Enabled = True
    
    pSetFocus tdbcLibro
    tdbtCodigo.ReadOnly = True
    
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
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
End Sub

Private Sub Imprimir()
    Dim matriz(15) As Variant
    Dim Titulo As String
    Titulo = "Tipos de Asiento"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;LIBRO - DESCRIPCION;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;;True"
    matriz(5) = "@Titulo04;;True"
    matriz(6) = "@Titulo05;DEBE/HABER;True"
    matriz(7) = "@Titulo06;CUENTA;True"
    matriz(8) = "@Titulo07;PORCENT.;True"
    matriz(9) = "@Tipo;TIPO_ASIENTO;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"
    
    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandarAgrupado.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub

Private Function ValidaDestinoPorc() As Boolean
 
    ValidaDestinoPorc = True
End Function

Private Sub Grabar()
    If validarDatos = False Then Exit Sub

    Dim clsMante As clsMantoTablas
    Dim i As Integer
    If lTipoMnt = "INSERTAR" Then Call GeneraCodigo
    
    Set clsMante = New clsMantoTablas
    On Local Error GoTo ErrorEjecucion
    Call CargaArregloMnt(0)
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaConceptosXLibros", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    Call Cancelar
    Call CargaTabla
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
    If CE(tdbtCodigo.Text) = "" Then
        Mensajes "Ingrese el codigo", vbInformation
        pSetFocus tdbtCodigo
        Exit Function
    End If
    
    If CE(tdbtDescripcion.Text) = "" Then
        Mensajes "Ingrese la descripcion", vbInformation
        pSetFocus tdbtDescripcion
        Exit Function
    End If
    
    
    validarDatos = True
End Function

Private Sub CargaArregloMnt(Numero As Integer)
    
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(9) As Variant
    lArrMnt(0) = lTipoMnt           ' Accion
    lArrMnt(1) = gsEmpresa          ' Empresa
    lArrMnt(2) = gsAnio         ' Año
    lArrMnt(3) = Me.tdbcLibro.BoundText    ' Libro
    lArrMnt(4) = Me.tdbtCodigo.Text                 ' Codigo
    lArrMnt(5) = Me.tdbtDescripcion          ' descripcion}
    lArrMnt(6) = CE(chkDetallado.Value)     ' detallado
    lArrMnt(7) = NE(chkDefecto.Value)     ' defecto
    lArrMnt(8) = "A"
    lArrMnt(9) = gsUsuario
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
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
                " " & _
                "ORDER BY LIB_CDESCRIPCION "
                '"LIB_CTIPOLIBRO='" & lsLibroCajIng & "' OR LIB_CTIPOLIBRO='" & lsLibroCajEgr & "' OR " & _
                '"LIB_CTIPOLIBRO='" & lsLibroDiario & "') " & _

    LlenarComboAddItem tdbcLibro, sqlcombos

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)

    Call pCargaCfgLibro
    Call LlenaCombos
    Call CargaTabla
    
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    SSTCentroCosto.Tab = 0
    
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    Call CerrarRecordSet(lrsTabla)
    
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    'Set lrstabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaConceptosXLibros 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '', ''"
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos)
    If Not lrsTabla Is Nothing Then
        ' *** Llenar grilla con el RecordSet
        lrsTabla.Sort = "Lib_cTipoLibro, Asl_cCodigo"
        tdbgCostos.DataSource = lrsTabla
        ' ***
        Exit Sub
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas

    Dim sqlSp As String
    sqlSp = "spCn_GrabaConceptosXLibros 'BUSCARREGISTRO', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbgCostos.Columns(1).Value & "', '" & tdbgCostos.Columns(2).Value & "'"
    
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro el registro. Probablemente eliminado desde otra sesion", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    
    tdbcLibro.BoundText = CE(rsArreglo!Lib_cTipoLibro)
    tdbtCodigo.Text = CE(rsArreglo!Asl_cCodigo)
    tdbtDescripcion.Text = CE(rsArreglo!Asl_cDescripcion)
    chkDetallado.Value = NE(rsArreglo!Asl_cDetallado)
    On Error Resume Next
    chkDefecto.Value = NE(rsArreglo!Asl_cDefecto)


End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(1) As String
    Dim i As Integer
    
    If lrsTabla Is Nothing Then Exit Sub
    On Local Error GoTo ErrorEjecucion
    cadena = ""
    If Trim(tdbtDescripcionBus) <> "" Then filtros(0) = "Asl_cDescripcion like '*" & tdbtDescripcionBus & "*'"
    If Trim(tdbtLibroBus) <> "" Then filtros(1) = "Lib_cDescripcion like '*" & tdbtLibroBus & "*'"
    For i = 0 To 1
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
    Exit Sub
ErrorEjecucion:
    If Err.Number <> "3265" Then Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub tdbcLibro_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbcLibro.Locked = True
    Else
        tdbcLibro.Locked = False
    End If
    ' ***
End Sub

Private Sub tdbcLibro_ItemChange()
    If lTipoMnt = "INSERTAR" Then Call GeneraCodigo

    If tdbcLibro.BoundText <> lsLibroDiario Then
        chkDefecto.Enabled = True
    Else
        chkDefecto.Enabled = False
        chkDefecto.Value = vbUnchecked
    End If
    
    
End Sub

Private Sub GeneraCodigo()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim lrsCodigo As New ADODB.Recordset
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaConceptosXLibros 'SEL_COR', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcLibro.BoundText & "'"
    arrDatos = Array(sqlSp)
    Set lrsCodigo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsCodigo Is Nothing Then
        ' *** Hallar Codigo
        Me.tdbtCodigo = lrsCodigo(0).Value
        Exit Sub
    End If
    
    Set lrsCodigo = Nothing
End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbcLibro_LostFocus()
    Call tdbcLibro_ItemChange
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

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

Private Sub tdbtLibroBus_Change()
    
    If gsKey = 219 Then
       tdbtLibroBus = Replace(tdbtLibroBus, "'", "")
       tdbtLibroBus.SelStart = Len(tdbtLibroBus)
    End If
    
    Call FiltrarRecordSet
End Sub

Private Sub tdbtLibroBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub
