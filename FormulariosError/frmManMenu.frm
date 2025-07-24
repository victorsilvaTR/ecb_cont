VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Acceso  al Sistema"
   ClientHeight    =   5610
   ClientLeft      =   3030
   ClientTop       =   2985
   ClientWidth     =   9795
   Icon            =   "frmManMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   9795
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5190
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9155
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ForeColor       =   14737632
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmManMenu.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4875
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   9345
         Begin TrueOleDBGrid70.TDBDropDown tdbdOperaTC 
            Height          =   1470
            Left            =   1080
            TabIndex        =   10
            Top             =   2700
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   1720
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Descripcion"
            Columns(0).DataField=   "pfl_cDescripcion"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "pfl_cCodPerfil"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=5080"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5001"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1164"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1085"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   -1  'True
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   16777215
            ValueTranslate  =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Arial"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=Arial"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=Arial"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.bgcolor=&HFFFFFF&"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(45)  =   "Named:id=33:Normal"
            _StyleDefs(46)  =   ":id=33,.parent=0"
            _StyleDefs(47)  =   "Named:id=34:Heading"
            _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(49)  =   ":id=34,.wraptext=-1"
            _StyleDefs(50)  =   "Named:id=35:Footing"
            _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   "Named:id=36:Selected"
            _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   2745
            Left            =   135
            TabIndex        =   6
            Top             =   2070
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   4842
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo "
            Columns(0).DataField=   "opm_cCodMenu"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "opm_cDesMenu"
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
            Columns(2).ValueItems(1).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
            Columns(2).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(2)=   "//////////////////////////////////////////////////8AMd4AMd7///8AMd4AMd7/////"
            Columns(2).ValueItems(1).DisplayValue(3)=   "//////////////////////////////////8AMd4AMd7///////8AMd4AMd4AMd7/////////////"
            Columns(2).ValueItems(1).DisplayValue(4)=   "//////////////////8AMd4AMd7///////////8AMd4AMd4AMd4AMd7/////////////////////"
            Columns(2).ValueItems(1).DisplayValue(5)=   "//8AMd4AMd7///////////////////8AMe8AMd4AMd4AMd7///////////////8AMd4AMd7/////"
            Columns(2).ValueItems(1).DisplayValue(6)=   "//////////////////////////8AMd4AMd4AMd7///8AMd4AMd4AMd7/////////////////////"
            Columns(2).ValueItems(1).DisplayValue(7)=   "//////////////////8AMd4AMecAMecAMecAMd7/////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(8)=   "//////////8AMecAMecAMe////////////////////////////////////////////////8AMd4A"
            Columns(2).ValueItems(1).DisplayValue(9)=   "Me8AMecAMe8AMff///////////////////////////////////////8AMfcAMe8AMef///////8A"
            Columns(2).ValueItems(1).DisplayValue(10)=   "MfcAMff///////////////////////////////8AMf8AMe8AMff///////////////8AMf8AMff/"
            Columns(2).ValueItems(1).DisplayValue(11)=   "//////////////////////8AMfcAMfcAMf////////////////////////8AMfcAMff/////////"
            Columns(2).ValueItems(1).DisplayValue(12)=   "//////8AMfcAMfcAMff///////////////////////////////////8AMff///////8AMfcAMfcA"
            Columns(2).ValueItems(1).DisplayValue(13)=   "Mff///////////////////////////////////////////////////8AMfcAMff/////////////"
            Columns(2).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////8="
            Columns(2).ValueItems(1).DisplayValue.vt=   9
            Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems.Count=   2
            Columns(2).Caption=   "Activo"
            Columns(2).DataField=   "opm_cActivado"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "OPM_CTITULO"
            Columns(3).DataField=   "OPM_CTITULO"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "OPM_CNOMOBJ"
            Columns(4).DataField=   "OPM_CNOMOBJ"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Privilegios"
            Columns(5).DataField=   "pfl_cDescripcion"
            Columns(5).DropDown=   "tdbdOperaTC"
            Columns(5).DropDown.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Codigo"
            Columns(6).DataField=   "PFL_CCODPERFIL"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1005"
            Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7382"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerStyle=3"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=7276"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=1164"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerStyle=3"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(22)=   "Column(3).Width=212"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=132"
            Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(29)=   "Column(4).Width=1217"
            Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1138"
            Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(34)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(36)=   "Column(5).Width=4657"
            Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=4577"
            Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=532"
            Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(42)=   "Column(5).DropDownList=1"
            Splits(0)._ColumnProps(43)=   "Column(5).AutoCompletion=1"
            Splits(0)._ColumnProps(44)=   "Column(6).Width=1244"
            Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=1164"
            Splits(0)._ColumnProps(47)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=532"
            Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
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
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=63,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(26)  =   ":id=63,.fgcolor=&H80000008&"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=72,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=64,.parent=2,.alignment=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=65,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=68,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=67,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=69,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=70,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=71,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=73,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=74,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=78,.parent=63,.bgcolor=&HFFFFFF&"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=64"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=65"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=67"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=82,.parent=63,.bgcolor=&HFFFFFF&"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=64"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=65"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=67"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=86,.parent=63,.bgcolor=&HFFFFFF&"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=64"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=65"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=67"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=90,.parent=63"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=64"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=65"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=67"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=94,.parent=63"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=64"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=65"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=67"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=98,.parent=63,.bgcolor=&HFFFFFF&"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=64"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=65"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=67"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=102,.parent=63,.bgcolor=&HFFFFFF&"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=64"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=65"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=67"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList70.TDBCombo tdbcUsuario 
            Height          =   300
            Left            =   2745
            TabIndex        =   9
            Tag             =   "_"
            Top             =   360
            Width           =   5130
            _ExtentX        =   9049
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2381"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2302"
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
            DataMode        =   4
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
            CellTips        =   1
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
            _PropDict       =   $"frmManMenu.frx":0EE6
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
         Begin TrueOleDBList70.TDBCombo tdbcPerfil 
            Height          =   300
            Left            =   2745
            TabIndex        =   1
            Tag             =   "_"
            Top             =   1170
            Width           =   5130
            _ExtentX        =   9049
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
            DataMode        =   4
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
            CellTips        =   1
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
            _PropDict       =   $"frmManMenu.frx":0F6D
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
         Begin TrueOleDBList70.TDBCombo tdbcEmpresa 
            Height          =   300
            Left            =   2745
            TabIndex        =   11
            Tag             =   "enabled"
            Top             =   765
            Width           =   5130
            _ExtentX        =   9049
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
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
            _PropDict       =   $"frmManMenu.frx":0FF4
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
         Begin MSForms.CommandButton cmdCargarPerfil 
            Height          =   390
            Left            =   2760
            TabIndex        =   2
            ToolTipText     =   " Cargar la config. del perfil a la lista "
            Top             =   1560
            Width           =   1575
            Caption         =   " Cargar Perfil"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManMenu.frx":107B
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
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
            Left            =   1575
            TabIndex        =   12
            Top             =   810
            Width           =   765
         End
         Begin MSForms.CommandButton cmdListar 
            Height          =   390
            Left            =   4545
            TabIndex        =   3
            ToolTipText     =   " Cargar la configuración guardada del usuario seleccionado"
            Top             =   1560
            Width           =   1575
            Caption         =   " Cargar Config."
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManMenu.frx":1615
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   390
            Left            =   6300
            TabIndex        =   4
            ToolTipText     =   " Actualiza accesos del sistema ( Solo si es el usuario que inicio sesion )"
            Top             =   1560
            Width           =   1575
            Caption         =   " Actualizar Menu"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManMenu.frx":1BAF
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Perfil"
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
            Left            =   1575
            TabIndex        =   8
            Top             =   1260
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
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
            Left            =   1575
            TabIndex        =   7
            Top             =   360
            Width           =   660
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
            Picture         =   "frmManMenu.frx":2149
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":2523
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":28FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":2CD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":30B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":348B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":3865
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":3C3F
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
            Picture         =   "frmManMenu.frx":4C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":4DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":4F0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":5067
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":51C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":531B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":55CF
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
            Picture         =   "frmManMenu.frx":5729
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":5CC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":625D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":67F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":6D91
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":732B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":78C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManMenu.frx":7E5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   13
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Editar F6"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "frmManMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim lArrMenu As New XArrayDB
Dim lrsOperaTC As New ADODB.Recordset
Dim inicio As Boolean
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Sub cmdCargarPerfil_Click()
    CargaPerfiles
End Sub

Private Sub cmdListar_Click()
    If CE(tdbcUsuario.BoundText) <> "" Then
        CargaTabla
    Else
        Mensajes "Seleccione un usuario"
    End If
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdRefresh_Click()
    If MsgBox("Se grabara la configuración del usuario seleccionado, desea continuar", vbQuestion + vbYesNo) = vbYes Then
        If CE(tdbcUsuario.BoundText) <> "" Then
            Grabar
            DoEvents
            frmMDIConta.CargaValoresMenu (CE(tdbcEmpresa.BoundText))
        Else
            Mensajes "Seleccione un usuario"
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()

On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 200
            .Height = Me.Height - .Top - 400
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 300
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 100
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 200
            .Height = Frame1.Height - .Top - 200
        End With

        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    
    Select Case Button.Index
        Case 1: 'ManNuevo
        Case 2: 'VerDatos
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                Me.tdbcUsuario.Enabled = True
        Case 4: 'Borrar
        Case 5: 'Editar
        Case 6: 'Imprimir
        Case 7
                respuesta = MsgBox("Esta seguro que desea salir del mantenimiento de accesos", vbYesNo + vbQuestion, "Confirmar Salir")
                If respuesta = vbYes Then Unload Me
                
    End Select
End Sub

Private Sub Grabar()

    If tdbcUsuario.BoundText = "" Then Mensajes "Seleccione un Usuario.": Call pSetFocus(tdbcUsuario): Exit Sub
    If tdbcEmpresa.BoundText = "" Then Mensajes "Seleccione una Empresa.": Call pSetFocus(tdbcEmpresa): Exit Sub
    
    Dim cn As New ADODB.Connection
    Dim sql As String
    
    Dim x As Integer
     x = 0
    cn.ConnectionString = gsCadenaConexion
    cn.Open
  Do While Not lrsTabla.EOF()
  
 
    On Local Error GoTo ErrorEjecucion
    
    If Not lrsTabla Is Nothing Then
        If Not (lrsTabla.EOF And lrsTabla.BOF) Then
            If lrsTabla.RecordCount > 0 Then
                
                sql = "spSg_GrabaAcceso 'EDITAR', '" & gsSOFT & "', '" & tdbcUsuario.BoundText & "'" _
                        & ", '" & lArrMenu(x, 1) & "', '" & CE(lArrMenu(x, 0)) & "', '" & tdbcEmpresa.BoundText & "', '" & CE(lArrMenu(x, 7)) & "'"
                
                'lArrMenu
                'CE(lrsTabla.Fields("Opm_cActivado")),lrsTabla.Fields("Opm_cCodMenu")
                cn.Execute (sql)
                
                
                x = x + 1
             Debug.Print sql
             
            End If
        End If
    End If
    
    lrsTabla.MoveNext
    
    Loop

    '----------------
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    
    Mensajes "Se grabaron los perfiles con exito"
    
    '*** ACTUALIZAR LAS OPCIONES DEL MENU SI LOS CAMBIOS SON PARA EL USUARIO ACTUAL
    If tdbcUsuario.BoundText = gsUsuario Then
    
        frmMDIConta.CargaValoresMenu (tdbcEmpresa.BoundText)
        Mensajes "Se actualizo el menu actual del sistema"
        
    End If

    Exit Sub
ErrorEjecucion:
    Mensajes "Error:" & Str(Err.Number) & Chr(13) & Err.Description & Chr(13) & "Source:" & Err.Source
    
End Sub

Private Sub CargaArregloMnt()
    ReDim lArrMnt(4) As Variant
    lArrMnt(0) = "EDITAR"
    lArrMnt(1) = "001"
    lArrMnt(2) = ""
    lArrMnt(3) = ""
    lArrMnt(4) = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(0) = True Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                'respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                'If respuesta = vbYes Then Call Cancelar
            End If
        'Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        'Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        'Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar
        'Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        'Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub
Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    
    inicio = True
    CargaUsuarios
    DoEvents
    CargaComboPerfiles
    DoEvents
    'CargaEmpresa
    'DoEvents
    'CargaTabla
    
    tdbcPerfil.Enabled = True
    
    lRegElim = False
    lTipoMnt = "INSERTAR"
    

    SSTCentroCosto.Tab = 0
    tdbgCostos.Columns(1).FetchStyle = True
    tbrOpciones.Buttons(3).Enabled = True
    tbrOpciones.Buttons(5).Visible = False
    
    'tdbcUsuario.Bookmark = 0
    'tdbcPerfil.Bookmark = 0
    'tdbcEmpresa.Bookmark = 0
    
    tdbcEmpresa.Columns(2).Visible = False
    tdbcEmpresa.Columns(0).Visible = False
    tdbcPerfil.Columns(0).Visible = False
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdCargarPerfil.Enabled = False
        cmdListar.Enabled = False
        cmdRefresh.Enabled = False
        
        tdbcPerfil.Locked = True
    Else
        cmdCargarPerfil.Enabled = True
        cmdListar.Enabled = True
        cmdRefresh.Enabled = True
        
        tdbcPerfil.Locked = False
    End If
    
    
    SSTCentroCosto.TabCaption(0) = ""
    'DoEvents
    'tdbcUsuario.BoundText = gsUsuario
    'DoEvents
    'tdbcUsuario_ItemChange
    'DoEvents
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
End Sub

Private Sub CargaEmpresa()
    Dim sqlCadena  As String
    Dim sAux As String

    Set tdbcEmpresa.DataSource = Nothing
    tdbcEmpresa.Text = ""
    
    sqlCadena = "spCN_GestionEmpresas 'BUSCARXUSUARIO', '','','','','','','','','','','','','','" & tdbcUsuario.BoundText & "'"
    
    LlenarComboAddItem tdbcEmpresa, sqlCadena
End Sub

Private Sub CargaComboPerfiles()
    Dim sqlSp As String
    'Dim lArr As New XArrayDB
    Dim lArrPerfil As New XArrayDB

    'LLENA COMBO DE PERFILES
    sqlSp = "exec spSg_GrabaUsuarios 'SEL_ALL_PERFILES', '','','','','','', '" & tdbcEmpresa.BoundText & "', '','" & gsSOFT & "'"
    ComboArreglo lArrPerfil, tdbcPerfil, sqlSp
    
    tdbcPerfil.Bookmark = 0
    tdbcPerfil.Enabled = True

End Sub

Private Sub CargaUsuarios()
    Dim sqlSp As String
    Dim lArr As New XArrayDB
    'Dim lArrPerfil As New XArrayDB
    
    'LLENA COMBO DE USUARIOS
    sqlSp = "EXEC spSg_GrabaAcceso 'SEL_USR','" & gsSOFT & "', '" & gsUsuario & "',  '',''"
    ComboArreglo lArr, tdbcUsuario, sqlSp
    
    'tdbcUsuario.BoundText = gsUsuario
    
    tdbcUsuario.Bookmark = 0
    tdbcUsuario.Enabled = True
    
    'LLENA COMBO DE PRIVILEGIOS
    sqlSp = "SELECT pfl_cCodPerfil, pfl_cDescripcion FROM SGM_PRIVILEGIOS ORDER BY pfl_cOrden"
    Call LlenarRecordSet(sqlSp, lrsOperaTC)
    Set Me.tdbdOperaTC.DataSource = lrsOperaTC
    tdbdOperaTC.Columns(1).Width = 0
    tdbdOperaTC.Columns(1).Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub CargaPerfiles()
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim rsPerfil As New ADODB.Recordset
    Dim i As Integer
    
    If tdbcPerfil.BoundText = "" Then
        Mensajes "Seleccione un Perfil"
        Exit Sub
    End If
    'Set tdbgCostos.DataSource = Nothing
    On Error GoTo ERROR
    If tdbcPerfil.BoundText <> "01" Then
        sqlSp = "spSg_GrabaAcceso 'SEL_ALL_PLANTILLA_PERFILES', '" & gsSOFT & "', '" & CE(tdbcPerfil.BoundText) & "', '','','" & CE(tdbcEmpresa.BoundText) & "'"
        arrDatos = Array(sqlSp)
        
        Set rsPerfil = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
        
        If Not rsPerfil Is Nothing Then
            If rsPerfil.State = adStateOpen Then
            For i = 0 To lArrMenu.Count(1) - 1
                rsPerfil.MoveFirst
                
                Do While Not rsPerfil.EOF
                    If CE(rsPerfil!OPM_Ccodmenu) = CE(lArrMenu(i, 0)) Then
                        lArrMenu(i, 1) = CE(rsPerfil!OPM_CACTIVADO)
                        tdbgCostos.Bookmark = i + 1
                        tdbgCostos.Columns(2) = CE(rsPerfil!OPM_CACTIVADO)
                        Exit Do
                    End If
                    rsPerfil.MoveNext
                Loop
            Next i
            Else
                Mensajes "No hay perfiles de tipo " & UCase(tdbcPerfil.Text)
            End If
        End If
    
    End If
    CerrarRecordSet rsPerfil
    Set clDatos = Nothing
    
    On Error Resume Next
    tdbgCostos.Bookmark = 1
    
    tdbgCostos.Refresh
    DoEvents
    tdbgCostos.ReBind
    Exit Sub
ERROR:
    Mensajes Err.Description, vbOKOnly + vbInformation
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    
    CerrarRecordSet lrsTabla
    tdbgCostos.DataSource = Nothing
    
    Set lrsTabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
        
    sqlSp = "spSg_GrabaAcceso 'SEL_ALL', '" & gsSOFT & "', '" & CE(tdbcUsuario.BoundText) & "', '','','" & CE(tdbcEmpresa.BoundText) & "'"
       
    'arrDatos = Array(sqlSp)
    'LlenarArreglo lArrMenu, sqlSp
    'Set tdbgCostos.Array = lArrMenu
    'Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    Set lrsTabla = fRetornaRS(sqlSp)
    Call LLenarArregloRS(lArrMenu, lrsTabla)
    Set tdbgCostos.DataSource = Nothing
    Set tdbgCostos.Array = lArrMenu
    
    If Not lrsTabla Is Nothing Then
        lrsTabla.Sort = "OPM_CCODMENU"
        tdbgCostos.DataSource = lrsTabla
        tdbgCostos.Columns(0).Visible = False
        tdbgCostos.Columns(3).Visible = False
        tdbgCostos.Columns(4).Visible = False
        tdbgCostos.Columns(6).Visible = False
        tdbgCostos.Columns(2).Alignment = dbgCenter
    End If
    
    
    'tdbcPerfil.Bookmark = 0

End Sub

Private Sub tdbcEmpresa_ItemChange()
CargaTabla
End Sub

Private Sub tdbcPerfil_ItemChange()
    If CE(tdbcPerfil.BoundText) = "01" Then
        cmdCargarPerfil.Enabled = False
    Else
        cmdCargarPerfil.Enabled = True
    End If

End Sub

Private Sub tdbcUsuario_ItemChange()
    
    'CargaUsuarios
     DoEvents
    CargaEmpresa

    DoEvents
    CargaComboPerfiles
    DoEvents
    CargaTabla
    
    If CE(tdbcUsuario.BoundText) = gsUsuario Then
        cmdRefresh.Enabled = True
    Else
        cmdRefresh.Enabled = False
    End If
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdRefresh.Enabled = False
    Else
        If CE(tdbcUsuario.BoundText) = gsUsuario Then
            cmdRefresh.Enabled = True
        Else
            cmdRefresh.Enabled = False
        End If
    End If
End Sub

Private Sub tdbdOperaTC_RowChange()
    tdbgCostos.Columns(6).Value = tdbdOperaTC.Columns(1).Value
    lArrMenu(tdbgCostos.Bookmark - 1, 7) = tdbdOperaTC.Columns(1).Value
End Sub

Private Sub tdbgCostos_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex <> 5 Or tbrOpciones.Buttons(3).Enabled = False Then
        Cancel = 1
    End If
End Sub

Private Sub tdbgCostos_DblClick()
    If tdbgCostos.Col = 2 And tbrOpciones.Buttons(3).Enabled = True Then
        If NE(tdbgCostos.Bookmark) > 0 Then
            If lArrMenu(tdbgCostos.Bookmark - 1, 1) = "0" Then
                lArrMenu(tdbgCostos.Bookmark - 1, 1) = "1"
                tdbgCostos.Columns(2) = "1"
            Else
                lArrMenu(tdbgCostos.Bookmark - 1, 1) = "0"
                tdbgCostos.Columns(2) = "0"
            End If
        End If
    End If
End Sub

Private Sub tdbgCostos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
    If Col > 0 Then
        If CE(lArrMenu(Bookmark - 1, 2)) = "S" Then
            CellStyle.ForeColor = &H800000
            CellStyle.Font.Bold = True
            
            If Left(lArrMenu(Bookmark - 1, 3), 1) <> " " Then
                CellStyle.Font.Underline = True
            End If
            
        Else
            CellStyle.ForeColor = &H80000012
            CellStyle.Font.Bold = False
            CellStyle.Font.Underline = False
            
        End If
        
        If tdbgCostos.Bookmark = Bookmark Then
            CellStyle.ForeColor = tdbgCostos.SelectedForeColor
            CellStyle.BackColor = tdbgCostos.SelectedBackColor
        End If
    End If
End Sub

