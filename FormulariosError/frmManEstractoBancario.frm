VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManEstractoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Extracto Bancario"
   ClientHeight    =   6330
   ClientLeft      =   495
   ClientTop       =   1755
   ClientWidth     =   10800
   Icon            =   "frmManEstractoBancario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10800
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   45
      TabIndex        =   7
      Top             =   810
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "MOVIMIENTO CONCILIADO"
      TabPicture(0)   =   "frmManEstractoBancario.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMoneda(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbnTC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdbtCambio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraMarco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraBotones"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdbgDatos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TDBDate1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "MOVIMIENTO SIN CONCILIAR"
      TabPicture(1)   =   "frmManEstractoBancario.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBotonesImp"
      Tab(1).Control(1)=   "tdbgDatosImp"
      Tab(1).Control(2)=   "lblMoneda(1)"
      Tab(1).ControlCount=   3
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   300
         Left            =   5760
         TabIndex        =   11
         Tag             =   "enabled"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         Calendar        =   "frmManEstractoBancario.frx":0F02
         Caption         =   "frmManEstractoBancario.frx":1004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmManEstractoBancario.frx":1068
         Keys            =   "frmManEstractoBancario.frx":1086
         Spin            =   "frmManEstractoBancario.frx":10F2
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   38974
         CenturyMode     =   0
      End
      Begin TrueOleDBGrid70.TDBGrid tdbgDatos 
         Height          =   3390
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   5980
         _LayoutType     =   4
         _RowHeight      =   15
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Mes"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Interno Voucher"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ItemVoucher"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Interno "
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   16
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   "I"
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "ABONO"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(1)._DefaultItem=   0
         Columns(4).ValueItems(1).Value=   "S"
         Columns(4).ValueItems(1).Value.vt=   8
         Columns(4).ValueItems(1).DisplayValue=   "CARGO"
         Columns(4).ValueItems(1).DisplayValue.vt=   8
         Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   2
         Columns(4).Caption=   "MOV"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TD"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DOCUMENTO"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "FECHA"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TC"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "External Editor"
         Columns(8).ExternalEditor=   "tdbnTC"
         Columns(8).ExternalEditor.vt=   8
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "MON. NAC."
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "MON. EXT."
         Columns(10).DataField=   ""
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "FECHA BANCO"
         Columns(11).DataField=   ""
         Columns(11).NumberFormat=   "External Editor"
         Columns(11).ExternalEditor=   "TDBDate1"
         Columns(11).ExternalEditor.vt=   8
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "OBSERVACION"
         Columns(12).DataField=   ""
         Columns(12).DataWidth=   250
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "GLOSA"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Editado"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "VOUCHER"
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Eliminado"
         Columns(16).DataField=   ""
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1905"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=212"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=132"
         Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(31)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(33)=   "Column(4).Width=1191"
         Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=1111"
         Splits(0)._ColumnProps(36)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(40)=   "Column(5).Width=953"
         Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=873"
         Splits(0)._ColumnProps(43)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(45)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(47)=   "Column(6).Width=2170"
         Splits(0)._ColumnProps(48)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(6)._WidthInPix=2090"
         Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(51)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(52)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(53)=   "Column(7).Width=1826"
         Splits(0)._ColumnProps(54)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(7)._WidthInPix=1746"
         Splits(0)._ColumnProps(56)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(57)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(58)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(59)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(60)=   "Column(8).Width=1349"
         Splits(0)._ColumnProps(61)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(8)._WidthInPix=1270"
         Splits(0)._ColumnProps(63)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(64)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(65)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(66)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(67)=   "Column(9).Width=2752"
         Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=2672"
         Splits(0)._ColumnProps(70)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(71)=   "Column(9)._ColStyle=514"
         Splits(0)._ColumnProps(72)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(73)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(74)=   "Column(10).Width=2514"
         Splits(0)._ColumnProps(75)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(10)._WidthInPix=2434"
         Splits(0)._ColumnProps(77)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(78)=   "Column(10)._ColStyle=514"
         Splits(0)._ColumnProps(79)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(80)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(81)=   "Column(11).Width=2355"
         Splits(0)._ColumnProps(82)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(11)._WidthInPix=2275"
         Splits(0)._ColumnProps(84)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=516"
         Splits(0)._ColumnProps(86)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(87)=   "Column(12).Width=6112"
         Splits(0)._ColumnProps(88)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(12)._WidthInPix=6033"
         Splits(0)._ColumnProps(90)=   "Column(12)._ColStyle=516"
         Splits(0)._ColumnProps(91)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(92)=   "Column(13).Width=6641"
         Splits(0)._ColumnProps(93)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(94)=   "Column(13)._WidthInPix=6562"
         Splits(0)._ColumnProps(95)=   "Column(13)._ColStyle=516"
         Splits(0)._ColumnProps(96)=   "Column(13).AllowFocus=0"
         Splits(0)._ColumnProps(97)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(98)=   "Column(14).Width=2328"
         Splits(0)._ColumnProps(99)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(100)=   "Column(14)._WidthInPix=2249"
         Splits(0)._ColumnProps(101)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(102)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(103)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(104)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(105)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(106)=   "Column(15).Width=1693"
         Splits(0)._ColumnProps(107)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(108)=   "Column(15)._WidthInPix=1614"
         Splits(0)._ColumnProps(109)=   "Column(15)._ColStyle=513"
         Splits(0)._ColumnProps(110)=   "Column(15).AllowFocus=0"
         Splits(0)._ColumnProps(111)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(112)=   "Column(16).Width=476"
         Splits(0)._ColumnProps(113)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(114)=   "Column(16)._WidthInPix=397"
         Splits(0)._ColumnProps(115)=   "Column(16).AllowSizing=0"
         Splits(0)._ColumnProps(116)=   "Column(16)._ColStyle=516"
         Splits(0)._ColumnProps(117)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(118)=   "Column(16).AllowFocus=0"
         Splits(0)._ColumnProps(119)=   "Column(16).Order=17"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   1
         CellTipsWidth   =   0
         DeadAreaBackColor=   14215660
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFBFFE1&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HF1EFEB&"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.bgcolor=&HFBFFE1&"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2,.bgcolor=&HFBFFE1&"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.bgcolor=&HFBFFE1&"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.bgcolor=&HFBFFE1&"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HFBFFE1&"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1,.bgcolor=&HFBFFE1&"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1,.bgcolor=&HFBFFE1&"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.bgcolor=&HFBFFE1&"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=98,.parent=13,.alignment=2,.bgcolor=&HFBFFE1&"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
         _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(105) =   "Named:id=33:Normal"
         _StyleDefs(106) =   ":id=33,.parent=0"
         _StyleDefs(107) =   "Named:id=34:Heading"
         _StyleDefs(108) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(109) =   ":id=34,.wraptext=-1"
         _StyleDefs(110) =   "Named:id=35:Footing"
         _StyleDefs(111) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(112) =   "Named:id=36:Selected"
         _StyleDefs(113) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(114) =   "Named:id=37:Caption"
         _StyleDefs(115) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(116) =   "Named:id=38:HighlightRow"
         _StyleDefs(117) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(118) =   "Named:id=39:EvenRow"
         _StyleDefs(119) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(120) =   "Named:id=40:OddRow"
         _StyleDefs(121) =   ":id=40,.parent=33"
         _StyleDefs(122) =   "Named:id=41:RecordSelector"
         _StyleDefs(123) =   ":id=41,.parent=34"
         _StyleDefs(124) =   "Named:id=42:FilterBar"
         _StyleDefs(125) =   ":id=42,.parent=33"
      End
      Begin VB.Frame fraBotonesImp 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   -74865
         TabIndex        =   24
         Top             =   4905
         Width           =   10410
         Begin MSForms.CommandButton cmdImportarTodos 
            Height          =   390
            Left            =   6075
            TabIndex        =   27
            ToolTipText     =   " Importa todos los movimientos "
            Top             =   45
            Width           =   1575
            Caption         =   "Importar Todos"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":111A
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdGrabaImp 
            Height          =   390
            Left            =   4455
            TabIndex        =   26
            ToolTipText     =   " Importa los movimientos seleccionados "
            Top             =   45
            Width           =   1575
            Caption         =   " Importar"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":16B4
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdImportar 
            Height          =   390
            Left            =   2835
            TabIndex        =   25
            ToolTipText     =   " Busca nuevos movimientos bancarios "
            Top             =   45
            Width           =   1575
            Caption         =   "Actualizar Lista"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":1C4E
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraBotones 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   180
         TabIndex        =   18
         Top             =   4905
         Width           =   10410
         Begin MSForms.CommandButton cmdGrabar 
            Height          =   390
            Left            =   7650
            TabIndex        =   23
            ToolTipText     =   " Graba los registros modificados "
            Top             =   45
            Width           =   1575
            Caption         =   " Grabar"
            PicturePosition =   327683
            Size            =   "2778;688"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEliminar 
            Height          =   390
            Left            =   4410
            TabIndex        =   22
            ToolTipText     =   " Elimina el regostro seleccionado "
            Top             =   45
            Width           =   1575
            Caption         =   " Eliminar"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":21E8
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdMostrar 
            Height          =   390
            Left            =   1170
            TabIndex        =   21
            ToolTipText     =   " Busca los movimientos conciliados "
            Top             =   45
            Width           =   1575
            Caption         =   "Actualizar Lista"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":2782
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEliminarTodo 
            Height          =   390
            Left            =   6030
            TabIndex        =   20
            ToolTipText     =   " Elimina todo el movimiento conciliado "
            Top             =   45
            Width           =   1575
            Caption         =   " Eliminar Todo"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":2D1C
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdVerificar 
            Height          =   390
            Left            =   2790
            TabIndex        =   19
            ToolTipText     =   " Verifica la existencia de los voucher registrados "
            Top             =   45
            Width           =   1575
            Caption         =   " Verificar"
            PicturePosition =   327683
            Size            =   "2778;688"
            Picture         =   "frmManEstractoBancario.frx":32B6
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraMarco 
         Height          =   780
         Left            =   180
         TabIndex        =   12
         Top             =   4095
         Width           =   10410
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "F11- Repetir Fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   4
            Left            =   8460
            TabIndex        =   17
            Top             =   165
            Width           =   1860
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   135
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Voucher Eliminados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   15
            Top             =   165
            Width           =   1680
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Voucher con Importes y/o Fechas modificadas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   13
            Top             =   435
            Width           =   3975
         End
      End
      Begin TDBNumber6Ctl.TDBNumber tdbtCambio 
         Height          =   330
         Left            =   7320
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         Calculator      =   "frmManEstractoBancario.frx":3850
         Caption         =   "frmManEstractoBancario.frx":3870
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmManEstractoBancario.frx":38D4
         Keys            =   "frmManEstractoBancario.frx":38F2
         Spin            =   "frmManEstractoBancario.frx":393C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,##0.000"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1802698757
         MinValueVT      =   1769209861
      End
      Begin TrueOleDBGrid70.TDBGrid tdbgDatosImp 
         Height          =   4035
         Left            =   -74865
         TabIndex        =   9
         Top             =   765
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   7117
         _LayoutType     =   4
         _RowHeight      =   15
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Mes"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Interno Voucher"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ItemVoucher"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Interno "
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   16
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   "I"
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "ABONO"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(1)._DefaultItem=   0
         Columns(4).ValueItems(1).Value=   "S"
         Columns(4).ValueItems(1).Value.vt=   8
         Columns(4).ValueItems(1).DisplayValue=   "CARGO"
         Columns(4).ValueItems(1).DisplayValue.vt=   8
         Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   2
         Columns(4).Caption=   "MOV"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TD"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DOCUMENTO"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "FECHA"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TC"
         Columns(8).DataField=   ""
         Columns(8).ExternalEditor=   "tdbtCambio"
         Columns(8).ExternalEditor.vt=   8
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "MON. NAC."
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "MON. EXT."
         Columns(10).DataField=   ""
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Fecha Banco"
         Columns(11).DataField=   ""
         Columns(11).NumberFormat=   "External Editor"
         Columns(11).ExternalEditor=   "TDBDate2"
         Columns(11).ExternalEditor.vt=   8
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Observacion"
         Columns(12).DataField=   ""
         Columns(12).DataWidth=   250
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "GLOSA"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Editado"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "VOUCHER"
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   16
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=16"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1905"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=212"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=132"
         Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(31)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(33)=   "Column(4).Width=1191"
         Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=1111"
         Splits(0)._ColumnProps(36)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(40)=   "Column(5).Width=953"
         Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=873"
         Splits(0)._ColumnProps(43)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(44)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(45)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(47)=   "Column(6).Width=2170"
         Splits(0)._ColumnProps(48)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(6)._WidthInPix=2090"
         Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(51)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(52)=   "Column(7).Width=1826"
         Splits(0)._ColumnProps(53)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(7)._WidthInPix=1746"
         Splits(0)._ColumnProps(55)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(56)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(57)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(58)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(59)=   "Column(8).Width=926"
         Splits(0)._ColumnProps(60)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(8)._WidthInPix=847"
         Splits(0)._ColumnProps(62)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(63)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(64)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(65)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(66)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(67)=   "Column(9).Width=2514"
         Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=2434"
         Splits(0)._ColumnProps(70)=   "Column(9)._ColStyle=514"
         Splits(0)._ColumnProps(71)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(72)=   "Column(10).Width=2328"
         Splits(0)._ColumnProps(73)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(74)=   "Column(10)._WidthInPix=2249"
         Splits(0)._ColumnProps(75)=   "Column(10)._ColStyle=514"
         Splits(0)._ColumnProps(76)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(77)=   "Column(11).Width=979"
         Splits(0)._ColumnProps(78)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(79)=   "Column(11)._WidthInPix=900"
         Splits(0)._ColumnProps(80)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(81)=   "Column(11)._ColStyle=516"
         Splits(0)._ColumnProps(82)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(83)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(84)=   "Column(12).Width=2090"
         Splits(0)._ColumnProps(85)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(86)=   "Column(12)._WidthInPix=2011"
         Splits(0)._ColumnProps(87)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(88)=   "Column(12)._ColStyle=516"
         Splits(0)._ColumnProps(89)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(90)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(91)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(92)=   "Column(13).Width=6641"
         Splits(0)._ColumnProps(93)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(94)=   "Column(13)._WidthInPix=6562"
         Splits(0)._ColumnProps(95)=   "Column(13)._ColStyle=516"
         Splits(0)._ColumnProps(96)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(97)=   "Column(14).Width=423"
         Splits(0)._ColumnProps(98)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(14)._WidthInPix=344"
         Splits(0)._ColumnProps(100)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(101)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(102)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(103)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(104)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(105)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(106)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(108)=   "Column(15)._ColStyle=513"
         Splits(0)._ColumnProps(109)=   "Column(15).Order=16"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   14215660
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE6ECE8&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HF1EFEB&"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1,.bgcolor=&HD7FCFF&"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1,.bgcolor=&HFBECEF&"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.bgcolor=&HFFFFFF&"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=98,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(101) =   "Named:id=33:Normal"
         _StyleDefs(102) =   ":id=33,.parent=0"
         _StyleDefs(103) =   "Named:id=34:Heading"
         _StyleDefs(104) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(105) =   ":id=34,.wraptext=-1"
         _StyleDefs(106) =   "Named:id=35:Footing"
         _StyleDefs(107) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(108) =   "Named:id=36:Selected"
         _StyleDefs(109) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(110) =   "Named:id=37:Caption"
         _StyleDefs(111) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(112) =   "Named:id=38:HighlightRow"
         _StyleDefs(113) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(114) =   "Named:id=39:EvenRow"
         _StyleDefs(115) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(116) =   "Named:id=40:OddRow"
         _StyleDefs(117) =   ":id=40,.parent=33"
         _StyleDefs(118) =   "Named:id=41:RecordSelector"
         _StyleDefs(119) =   ":id=41,.parent=34"
         _StyleDefs(120) =   "Named:id=42:FilterBar"
         _StyleDefs(121) =   ":id=42,.parent=33"
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnTC 
         Height          =   285
         Left            =   2025
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   570
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   503
         Calculator      =   "frmManEstractoBancario.frx":3964
         Caption         =   "frmManEstractoBancario.frx":3984
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmManEstractoBancario.frx":39E8
         Keys            =   "frmManEstractoBancario.frx":3A06
         Spin            =   "frmManEstractoBancario.frx":3A40
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16515041
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0.000"
         EditMode        =   2
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.000"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   6750213
         MinValueVT      =   3538949
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA DE LA CUENTA: NUEVOS SOLES"
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
         Height          =   285
         Index           =   1
         Left            =   -75000
         TabIndex        =   29
         Top             =   450
         Width           =   10680
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA DE LA CUENTA: NUEVOS SOLES"
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
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   450
         Width           =   10680
      End
   End
   Begin VB.Frame fraListas 
      Height          =   795
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   10710
      Begin TrueOleDBList70.TDBCombo tdbcBancoBus 
         Height          =   345
         Left            =   900
         TabIndex        =   0
         Tag             =   "_"
         Top             =   225
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   609
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   7938
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
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   345.26
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
         _PropDict       =   $"frmManEstractoBancario.frx":3A68
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
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
      Begin TrueOleDBList70.TDBCombo tdbcCuentaBus 
         Height          =   345
         Left            =   4635
         TabIndex        =   1
         Tag             =   "_"
         Top             =   225
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   609
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
         Columns(1).Caption=   "Cuenta"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Moneda"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cta Contable"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2143"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2064"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=741"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=661"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2619"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=2275"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=2196"
         Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(34)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   5
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   345.26
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
         _PropDict       =   $"frmManEstractoBancario.frx":3AEF
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
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
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   8640
         TabIndex        =   2
         Top             =   225
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
         _PropDict       =   $"frmManEstractoBancario.frx":3B76
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MES"
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
         Left            =   8145
         TabIndex        =   6
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BANCO"
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
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CUENTA"
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
         Left            =   3825
         TabIndex        =   4
         Top             =   270
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmManEstractoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' Creado por    :  Miguel Angel Lopez Sanabria
' Modificado    :  Miguel Angel Lopez Sanabria
' Descripción   :  Realiza registro de asientos contables. Asi mismo consultas
'                  simples de los asientos.
' Fecha Crea    :  13/09/2004
' Fecha Modi    :  13/09/2004
' -----------------------------------------------------------------------------

Option Explicit

Dim lArrDetalle As New XArrayDB
Dim lArrDetalleImp As New XArrayDB

Dim lArrDet() As Variant
Dim lrsTabla As New ADODB.Recordset

Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub Importar(Todos As Boolean)
    CargaTabla
    DoEvents

    If IsNull(tdbgDatos.Bookmark) Then
        Set tdbgDatos.DataSource = Nothing
        tdbgDatos.ReBind
    End If

    Dim i As Integer, j As Integer
    Dim Filas As Integer
    Dim FilasImp As Integer
    Dim entro As Boolean
    Dim FilasEliminar() As Integer
    
    entro = False
    FilasImp = lArrDetalleImp.Count(1) - 1
    j = 0
    
    For i = 0 To FilasImp
        If tdbgDatosImp.IsSelected(i) >= 0 Or Todos = True Then
            ReDim Preserve FilasEliminar(j + 1)
            
            FilasEliminar(j) = i + 1
            
            Filas = lArrDetalle.Count(1)
            
            If CE(lArrDetalleImp(i, 0)) <> "" Then
                entro = True
            End If
            
            lArrDetalle.ReDim 0, Filas, 0, 16
            
            lArrDetalle(Filas, 0) = CE(lArrDetalleImp(i, 0))
            lArrDetalle(Filas, 1) = CE(lArrDetalleImp(i, 1))
            lArrDetalle(Filas, 2) = CE(lArrDetalleImp(i, 2))
            lArrDetalle(Filas, 3) = CE(lArrDetalleImp(i, 3))
            lArrDetalle(Filas, 4) = CE(lArrDetalleImp(i, 4))
            lArrDetalle(Filas, 5) = CE(lArrDetalleImp(i, 5))
            lArrDetalle(Filas, 6) = CE(lArrDetalleImp(i, 6))
            lArrDetalle(Filas, 7) = CE(lArrDetalleImp(i, 7))
            lArrDetalle(Filas, 8) = CE(lArrDetalleImp(i, 8))
            lArrDetalle(Filas, 9) = CE(lArrDetalleImp(i, 9))
            lArrDetalle(Filas, 10) = CE(lArrDetalleImp(i, 10))
            lArrDetalle(Filas, 11) = "" 'CE(lArrDetalleImp(i, 11))
            lArrDetalle(Filas, 12) = CE(lArrDetalleImp(i, 12))
            lArrDetalle(Filas, 13) = CE(lArrDetalleImp(i, 13))
            lArrDetalle(Filas, 14) = "1" 'CAMPO VERIFICADOR DE INSERTADO
            lArrDetalle(Filas, 15) = CE(lArrDetalleImp(i, 15))
            lArrDetalle(Filas, 16) = ""
            j = j + 1
        End If
    Next i
   
    If entro = True Then
        For i = j To 0 Step -1
            If FilasEliminar(i) >= 1 Then
                lArrDetalleImp.DeleteRows FilasEliminar(i) - 1
            End If
        Next i
        Set tdbgDatos.Array = lArrDetalle
        tdbgDatos.ReBind
        tdbgDatosImp.ReBind
    Else
        Mensajes "Seleccione minimo un movimiento de la lista para la importación", vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdEliminarTodo_Click()
        If ValidaCampos = False Then Exit Sub
        Eliminar (True)
End Sub

Private Sub cmdGrabaImp_Click()
    If ValidaCampos = False Then Exit Sub
    Importar (False)
End Sub

Private Sub cmdImportarTodos_Click()
    If ValidaCampos = False Then Exit Sub
    If MsgBox("Deseas importar todos los movimientos importados", vbYesNo + vbInformation, "Importando movimientos...") = vbYes Then
        Call Importar(True)
        SSTab1.Tab = 0
        
    End If
End Sub

Private Function ValidaCampos() As Boolean
    DoEvents
    If CE(tdbcBancoBus.Text) = "" Then
        Mensajes "Seleccione un banco de la lista", vbOKOnly + vbInformation
        ValidaCampos = False
        Exit Function
    End If
    
    If CE(tdbcCuentaBus.Text) = "" Then
        Mensajes "Seleccione una cuenta corriente de la lista", vbOKOnly + vbInformation
        ValidaCampos = False
        Exit Function
    End If
    
    ValidaCampos = True
End Function

Private Sub cmdGrabar_Click()
    If ValidaCampos = False Then Exit Sub
    If IsNull(tdbgDatos.Bookmark) Then Exit Sub
    Dim entro As Boolean
    entro = False
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    ' *** Validar q mes no haya sido cerrado
    If CierreMes(Me.tdbcMes.BoundText) = True Then
        Mensajes "El mes seleccionado ha sido cerrado. No se puede grabar.", vbInformation
        Exit Sub
    End If
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    
    tdbgDatos.Update
    
    DoEvents

    clsMante.InicializaClase
    clsMante.BeginTrans
    
    For i = 0 To lArrDetalle.Count(1) - 1
        If lArrDetalle(i, 14) = "1" Then    ' *** Grabar solo si se ha modificado
            Call CargaArregloDet(i)
            
            If CE(lArrDet(4)) = "" Then 'NUMERO INTERNO
                lArrDet(0) = "INSERTAR_AUTOM"
            End If
            
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaMovCheque", lArrDet(), False) = False Then
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                clsMante.CancelTrans
                clsMante.FinalizaClase
                Screen.MousePointer = vbNormal
                entro = False
                Exit Sub
            End If
            
            entro = True
        End If
    Next
    
    clsMante.CommitTrans
    clsMante.FinalizaClase
    
    
    
    If entro = True Then
    
        For i = 0 To lArrDetalle.Count(1) - 1
            lArrDetalle(i, 14) = ""
        Next i
        
        CargaTabla
        tdbgDatos.Refresh
        tdbgDatos.Update
        DoEvents
        Mensajes "Se ha grabado correctamente los movimientos conciliados.", vbInformation + vbOKOnly
            
        tdbgDatos.Refresh
        
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CargaArregloDet(item As Integer)
    ReDim lArrDet(21) As Variant
    lArrDet(0) = "EDITAR"              ' Empresa
    lArrDet(1) = gsEmpresa             ' Año
    lArrDet(2) = gsAnio
    lArrDet(3) = lArrDetalle(item, 0)  ' *** Periodo
    lArrDet(4) = lArrDetalle(item, 3)  ' *** Interno
    lArrDet(5) = lArrDetalle(item, 1)  ' *** InternoAsiento
    lArrDet(6) = lArrDetalle(item, 2)  ' *** Item Asiento
    lArrDet(7) = tdbcBancoBus.BoundText
    lArrDet(8) = tdbcCuentaBus.BoundText
    lArrDet(9) = lArrDetalle(item, 4)  ' *** TipoMov
    lArrDet(10) = lArrDetalle(item, 5) ' *** TipoDoc
    lArrDet(11) = lArrDetalle(item, 6) ' *** Numero Cheque
    
    ' *** Fecha de Cheque
    If lArrDetalle(item, 7) = "  /  /    " Then lArrDetalle(item, 7) = Null
    If lArrDetalle(item, 7) = "" Then lArrDetalle(item, 7) = Null
    lArrDet(12) = lArrDetalle(item, 7) ' *** Fecha Cheque
'    If lArrDetalle(item, 7) = "01/01/1900" Then lArrDetalle(item, 7) = Null
'    If lArrDetalle(item, 7) = "  /  /    " Then lArrDetalle(item, 7) = Null
'    If lArrDetalle(item, 7) = "" Then lArrDetalle(item, 7) = Null
    
    lArrDet(13) = lArrDetalle(item, 8) ' *** Tipo Cambio
    lArrDet(14) = lArrDetalle(item, 9) ' *** Soles
    lArrDet(15) = lArrDetalle(item, 10)    ' *** Dolares
    
    ' *** Fecha de Banco
    If lArrDetalle(item, 11) = "  /  /    " Then lArrDetalle(item, 11) = Null
    If lArrDetalle(item, 11) = "" Then lArrDetalle(item, 11) = Null
    lArrDet(16) = lArrDetalle(item, 11)    ' *** Fecha Banco
    
    lArrDet(17) = lArrDetalle(item, 12)    ' *** Observaciones
    lArrDet(18) = lArrDetalle(item, 13)    ' *** Glosa
    lArrDet(19) = "A"                   ' *** Estado
    lArrDet(20) = gsUsuario             ' *** Usuario
    lArrDet(21) = lArrDetalle(item, 15)             ' *** voucher
    
End Sub

Private Sub cmdImportar_Click()
    If ValidaCampos = False Then Exit Sub
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    Dim lrsImportados As New ADODB.Recordset
    Dim arrDatos() As Variant
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    sqlSp = "spCn_ConsultarAsientosEstBan '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & tdbcCuentaBus.Columns(4) & "','" & gsUsuario & "','" & tdbcBancoBus.BoundText & "','" & tdbcCuentaBus.BoundText & "' "
    arrDatos = Array(sqlSp)
    
    lArrDetalleImp.Clear
    Set tdbgDatosImp.Array = lArrDetalleImp
    
    Set lrsImportados = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsImportados.State = 1 Then
        lArrDetalleImp.ReDim 0, 0, 0, 16
        
        i = 0
        lrsImportados.MoveFirst
        Do While Not lrsImportados.EOF
            If i > 0 Then lArrDetalleImp.ReDim 0, i + 1, 0, 16
            i = i + 1
            
            lArrDetalleImp(i - 1, 0) = lrsImportados("Per_cPeriodo").Value
            lArrDetalleImp(i - 1, 1) = lrsImportados("Ase_cNummov").Value
            lArrDetalleImp(i - 1, 2) = lrsImportados("Asd_nItem").Value
            lArrDetalleImp(i - 1, 3) = "" 'lrsImportados("Che_nVoucherPago").Value
            lArrDetalleImp(i - 1, 4) = lrsImportados("TipoMov").Value
            lArrDetalleImp(i - 1, 5) = lrsImportados("Asd_cTipoDoc").Value
            lArrDetalleImp(i - 1, 6) = lrsImportados("Asd_cNumDoc").Value
            lArrDetalleImp(i - 1, 7) = FechaRegistro(lrsImportados("Asd_dFecDoc").Value)
            lArrDetalleImp(i - 1, 8) = lrsImportados("Asd_nTipoCambio").Value
            lArrDetalleImp(i - 1, 9) = lrsImportados("montoSoles").Value
            lArrDetalleImp(i - 1, 10) = lrsImportados("montodolar").Value
            lArrDetalleImp(i - 1, 11) = FechaRegistro(lrsImportados("Ase_dFecha").Value)
            lArrDetalleImp(i - 1, 12) = "" 'lrsImportados("Che_cObservacion").Value
            lArrDetalleImp(i - 1, 13) = lrsImportados("Asd_cGlosa").Value
            lArrDetalleImp(i - 1, 14) = ""
            lArrDetalleImp(i - 1, 15) = lrsImportados("Ase_nVoucher").Value
            lArrDetalleImp(i - 1, 16) = ""
            lrsImportados.MoveNext
        Loop
        
        lArrDetalleImp.ReDim 0, i - 1, 0, 16
    End If
    
    Set tdbgDatosImp.Array = lArrDetalleImp
    tdbgDatosImp.ReBind
    tdbgDatosImp.Bookmark = 0
    
    CerrarRecordSet lrsImportados
'    Me.tdbgDatosImp.Columns(13).Width = 3500
'    Me.tdbgDatosImp.Columns(15).Width = 1544
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabaImp.Enabled = False
    Else
        cmdGrabaImp.Enabled = True
    End If
    
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim i As Integer

    'lArrDetalle.Clear
    
    Set lArrDetalle = Nothing
    Set tdbgDatos.DataSource = Nothing
    Set tdbgDatos.Array = lArrDetalle
    
'    sqlSp = "spCn_ConsultarAsientosEstBan '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', " & _
'            "'" & tdbcCuentaBus.Columns(4) & "','" & gsUsuario & "','" & tdbcBancoBus.BoundText & "'," & _
'            "'" & tdbcCuentaBus.BoundText & "' "
            
    sqlSp = "spCn_GrabaMovCheque 'SEL_ALL_CTA', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '', " & _
            "'', '', '" & Trim(tdbcBancoBus.BoundText) & "', '" & tdbcCuentaBus.BoundText & "'"
    arrDatos = Array(sqlSp)
            
    arrDatos = Array(sqlSp)
    
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsTabla.State = 1 Then
    
        ' *** Llenar grilla con el RecordSet
        'lrsTabla.Sort = "Che_dFechaCheque, Che_cNummovVoucher, Che_nItemVoucher"
        lArrDetalle.ReDim 0, 0, 0, 16
        
        i = 0
        lrsTabla.MoveFirst
        ' *** Llenando el Arreglo
        Do While Not lrsTabla.EOF
            If i > 0 Then lArrDetalle.ReDim 0, i + 1, 0, 16
            i = i + 1
            
            lArrDetalle(i - 1, 0) = lrsTabla("Per_cPeriodo").Value
            lArrDetalle(i - 1, 1) = lrsTabla("Che_cNummovVoucher").Value
            lArrDetalle(i - 1, 2) = lrsTabla("Che_nItemVoucher").Value
            lArrDetalle(i - 1, 3) = lrsTabla("Che_nVoucherPago").Value
            lArrDetalle(i - 1, 4) = lrsTabla("Che_cTipoMov").Value
            lArrDetalle(i - 1, 5) = lrsTabla("Che_cTipoDoc").Value
            lArrDetalle(i - 1, 6) = lrsTabla("Che_cOperaCheque").Value
            lArrDetalle(i - 1, 7) = FechaRegistro(lrsTabla("Che_dFechaCheque").Value)
            lArrDetalle(i - 1, 8) = lrsTabla("Che_nTipoCambio").Value
            lArrDetalle(i - 1, 9) = lrsTabla("Che_nMontoS").Value
            lArrDetalle(i - 1, 10) = lrsTabla("Che_nMontoD").Value
            lArrDetalle(i - 1, 11) = FechaRegistro(lrsTabla("Che_dFechaOpera").Value)
            lArrDetalle(i - 1, 12) = lrsTabla("Che_cObservacion").Value
            lArrDetalle(i - 1, 13) = lrsTabla("Che_cGlosa").Value
            lArrDetalle(i - 1, 14) = "" 'CAMPO VERIFICADOR DE MODIFICACION
            lArrDetalle(i - 1, 15) = lrsTabla("Ase_nVoucher").Value
            lArrDetalle(i - 1, 16) = ""
            lrsTabla.MoveNext
        Loop
        
        ' *** Redimensionando arreglo y cerrando el recordSet
        lArrDetalle.ReDim 0, i - 1, 0, 16
    Else
        lArrDetalle.ReDim 0, 0, 0, 16
        lArrDetalle.DeleteRows (0)
    End If
    
    Set tdbgDatos.Array = lArrDetalle
    Set clDatos = Nothing
    'CerrarRecordSet lrsTabla
    tdbgDatos.Refresh
    tdbgDatos.ReBind
    tdbgDatos.Bookmark = 0
End Sub


Private Sub cmdMostrar_Click()
    If ValidaCampos = False Then Exit Sub
    CargaTabla
    
End Sub

Private Sub VerificarVoucher()
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim Msg1 As Boolean, Msg2 As Boolean, Msg3 As Boolean
    Dim i As Integer
    
    Msg1 = False
    Msg2 = False
    Msg3 = False

    sqlSp = "spCn_VerificarAsientosEstBan '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', " & _
            "'" & tdbcCuentaBus.Columns(4) & "','" & gsUsuario & "','" & tdbcBancoBus.BoundText & "'," & _
            "'" & tdbcCuentaBus.BoundText & "' "
    
    arrDatos = Array(sqlSp)
    
    For i = 0 To lArrDetalle.Count(1) - 1
        lArrDetalle(i, 16) = ""
    Next i
    
    Dim entro As Boolean
    
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsTabla.State = 1 Then
        
        For i = 0 To lArrDetalle.Count(1) - 1
            lArrDetalle(i, 14) = ""
        Next i
        entro = False
        
        Do While Not lrsTabla.EOF
            
            For i = 0 To lArrDetalle.Count(1) - 1
                If CE(lArrDetalle(i, 0)) = CE(lrsTabla.Fields("Per_cPeriodo")) And _
                   CE(lArrDetalle(i, 1)) = CE(lrsTabla.Fields("Ase_cNummov")) And _
                   CE(lArrDetalle(i, 15)) = CE(lrsTabla.Fields("Ase_nVoucher")) Then
                   lArrDetalle(i, 16) = "1"
                   entro = True
                   Exit For
                 End If
            Next i
            lrsTabla.MoveNext
        Loop
        If entro = True Then
            Mensajes "Se encontraron vouchers ELIMINADOS en los movimientos conciliados.", vbOKOnly + vbInformation
            Msg1 = True
        End If
    Else
        Msg1 = False
    End If
    
    '*****************************************************************************
    sqlSp = "spCn_VerificarImportesEstBan '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & tdbcCuentaBus.Columns(4) & "','" & gsUsuario & "','" & tdbcBancoBus.BoundText & "','" & tdbcCuentaBus.BoundText & "' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsTabla.State = 1 Then
        Do While Not lrsTabla.EOF
            
            For i = 0 To lArrDetalle.Count(1) - 1
                If CE(lArrDetalle(i, 0)) = CE(lrsTabla!Per_cPeriodo) And _
                   CE(lArrDetalle(i, 1)) = CE(lrsTabla!Ase_cNummov) And _
                   CE(lArrDetalle(i, 15)) = CE(lrsTabla!Ase_nVoucher) Then
                   lArrDetalle(i, 16) = "2"
                   Exit For
                 End If
            Next i
            lrsTabla.MoveNext
        Loop
        
        Mensajes "Se encontraron vouchers con IMPORTES modificados.", vbOKOnly + vbInformation
        Msg2 = True
    Else
        Msg2 = False
    End If
    '*****************************************************************************
    sqlSp = "spCn_VerificarFechasEstBan '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & tdbcCuentaBus.Columns(4) & "','" & gsUsuario & "','" & tdbcBancoBus.BoundText & "','" & tdbcCuentaBus.BoundText & "' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsTabla.State = 1 Then
        Do While Not lrsTabla.EOF
            
            For i = 0 To lArrDetalle.Count(1) - 1
                If CE(lArrDetalle(i, 0)) = CE(lrsTabla!Per_cPeriodo) And _
                   CE(lArrDetalle(i, 1)) = CE(lrsTabla!Ase_cNummov) And _
                   CE(lArrDetalle(i, 15)) = CE(lrsTabla!Ase_nVoucher) Then
                   lArrDetalle(i, 16) = "2"
                   Exit For
                 End If
            Next i
            lrsTabla.MoveNext
        Loop
        
        Mensajes "Se encontraron vouchers con FECHAS modificadas.", vbOKOnly + vbInformation
        Msg3 = True
    Else
        Msg3 = False
    End If
    '*****************************************************************************
    If Msg1 = False And Msg2 = False And Msg3 = False Then
        Mensajes "No se encontraron inconsistencias.", vbOKOnly + vbInformation
    End If
    
    Set lrsTabla = Nothing
    Set clDatos = Nothing
    
    tdbgDatos.Refresh
End Sub

Private Sub cmdVerificar_Click()
    If ValidaCampos = False Then Exit Sub
    VerificarVoucher
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    Call Centrar_form(Me)
    
    tdbgDatos.FetchRowStyle = True
    Dim sqlcombos As String
    Dim registros As Integer
    
    ' *** Llenando los Bancos
    sqlcombos = "select Ban_cCodigo, Ban_cNombre from CNT_BANCO " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' ORDER BY Ban_cNombre"
    registros = LlenarComboAddItem(tdbcBancoBus, sqlcombos)
    
    If registros <= 0 Then
        Mensajes "No se crearon bancos para esta empresa, agreguelos en configuración de bancos", vbOKOnly + vbInformation
    End If
    
    tdbcCuentaBus.DropdownWidth = tdbcCuentaBus.Width + tdbcCuentaBus.Width / 3
    Call LlenaComboMesApeAddItem(tdbcMes)
    
    If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
    
    
    tdbgDatosImp.Columns(11).Visible = False
    tdbgDatosImp.Columns(11).Width = 0
    tdbgDatosImp.Columns(12).Visible = False
    tdbgDatosImp.Columns(12).Width = 0
    Me.SSTab1.Tab = 0
    Me.cmdGrabar.Enabled = True
    tdbgDatos.Columns(16).FetchStyle = True
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdEliminar.Enabled = False
        cmdEliminarTodo.Enabled = False
        cmdGrabar.Enabled = False
        cmdGrabaImp.Enabled = False
        cmdImportarTodos.Enabled = False
    Else
        cmdEliminar.Enabled = True
        cmdEliminarTodo.Enabled = True
        cmdGrabar.Enabled = True
        cmdGrabaImp.Enabled = True
        cmdImportarTodos.Enabled = True
    
    End If
    
'    If gsPeriodo = "00" Then
'        tdbcMes.BoundText = "01"
'    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
'        tdbcMes.BoundText = "12"
'    Else
'        tdbcMes.BoundText = gsPeriodo
'    End If
    
    tdbcMes.ReBind
End Sub

Private Sub Form_Resize()
On Error GoTo serror
    If Me.WindowState <> vbMinimized Then
        SSTab1.Width = Me.Width - 200
        tdbgDatos.Width = SSTab1.Width - 300
        tdbgDatosImp.Width = SSTab1.Width - 300
        fraListas.Width = Me.Width - 200
        fraMarco.Width = Me.Width - 500
        fraBotones.Width = Me.Width - 500
        fraBotonesImp.Width = Me.Width - 500
        
        SSTab1.Height = Me.Height - fraListas.Height - 550
        fraMarco.Top = SSTab1.Top + SSTab1.Height - fraMarco.Height - 1500
        
        fraBotones.Top = fraMarco.Top + fraMarco.Height + 100
        fraBotonesImp.Top = fraBotones.Top
        
        tdbgDatos.Height = fraMarco.Top - tdbgDatos.Top
        tdbgDatosImp.Height = fraBotonesImp.Top - tdbgDatosImp.Top
        
        lblMONEDA(0).Left = 0
        lblMONEDA(0).Width = SSTab1.Width
        lblMONEDA(1).Left = 0
        lblMONEDA(1).Width = SSTab1.Width
    End If
serror:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lArrDetalle = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        cmdImportar_Click
    End If
    
End Sub

Private Sub tdbcBancoBus_ItemChange()
    Dim sqlcombos As String
    tdbcCuentaBus.Text = ""
    tdbcCuentaBus.Clear
    tdbcCuentaBus.Refresh
    tdbcCuentaBus.ReBind
    ' *** Llenando los Bancos
'    sqlcombos = "SELECT Cue_cNumCuenta as 'cuenta' , Cue_cNumCuenta,  CNM_CUENTA_BANCO.Mon_cCodigo, Mon_cNombreLargo, cue_cCuentaContable, Mon_cMNac " & _
'                "FROM CNM_CUENTA_BANCO LEFT JOIN CNT_TIPO_MONEDA " & _
'                "ON CNM_CUENTA_BANCO.Mon_cCodigo = CNT_TIPO_MONEDA.Mon_cCodigo " & _
'                "AND CNM_CUENTA_BANCO.Emp_cCodigo = CNT_TIPO_MONEDA.Emp_cCodigo " & _
'                "WHERE CNM_CUENTA_BANCO.Emp_cCodigo = '" & gsEmpresa & "' AND Ban_cCodigo = '" & tdbcBancoBus.BoundText & "' " & _
'                "ORDER BY CNM_CUENTA_BANCO.Mon_cCodigo, Cue_cNumCuenta "
    sqlcombos = "spCn_GrabaMovCheque 'BUSCACUENTABANCO','" & gsEmpresa & "','" & gsAnio & "','','','','','" & tdbcBancoBus.BoundText & "' "
    
    LlenarComboAddItem tdbcCuentaBus, sqlcombos
    ' ***
    cmdImportar_Click
End Sub

Private Sub tdbcBancoBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcCuentaBus
End If
End Sub

Private Sub tdbcCuentaBus_ItemChange()
'    If Me.tdbcCuentaBus.Columns(5) = "1" Then
'        Me.tdbgDatos.Splits(0).Columns(9).Visible = True
'        Me.tdbgDatos.Splits(0).Columns(10).Visible = False
'        Me.tdbgDatos.Splits(0).Columns(8).Visible = False
'    Else
'        Me.tdbgDatos.Splits(0).Columns(8).Visible = True
'        Me.tdbgDatos.Splits(0).Columns(9).Visible = False
'        Me.tdbgDatos.Splits(0).Columns(10).Visible = True
'    End If
    lblMONEDA(0).Caption = "MONEDA DE LA CUENTA: " & CE(tdbcCuentaBus.Columns(3).Value)
    lblMONEDA(1).Caption = "MONEDA DE LA CUENTA: " & CE(tdbcCuentaBus.Columns(3).Value)
    
    tdbgDatos.Columns(9).Caption = gsNombreMonedaNac
    tdbgDatos.Columns(10).Caption = gsNombreMonedaExt
    
    
    tdbgDatosImp.Columns(9).Caption = gsNombreMonedaNac
    tdbgDatosImp.Columns(10).Caption = gsNombreMonedaExt
    
    
    CargaTabla
    
    
    If SSTab1.Tab = 1 Then
        cmdImportar_Click
    End If
    
End Sub

Private Sub tdbcCuentaBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMes
End If
End Sub

Private Sub tdbcMes_ItemChange()
    Call tdbcCuentaBus_ItemChange
    
    If SSTab1.Tab = 1 Then
        cmdImportar_Click
    End If
    
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdMostrar
End If
End Sub

Private Sub TDBDate1_LostFocus()
    If Not IsNull(TDBDate1) Then
        If Format(TDBDate1.Value, "yyyyMMdd") < Format(tdbgDatos.Columns(7).Value, "yyyyMMdd") Then
            Mensajes "La fecha debe ser la misma o posterior a la ingresada", vbOKOnly + vbInformation
            TDBDate1.Value = ""
            tdbgDatos.Columns(11).Value = "__/__/____"
            tdbgDatos.Col = 11
            pSetFocus tdbgDatos
            
        Else
            On Error Resume Next
            lArrDetalle(tdbgDatos.Bookmark, 14) = "1"
            'tdbgDatos.Columns(14).Value = "1"
        End If
    End If
End Sub

Private Sub tdbgDatos_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 12 Then
        tdbgDatos.Columns(14).Value = "1"
    End If
End Sub

Private Sub tdbgDatos_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Cancel = 1
        Mensajes "No tiene privilegios de modificacion, contactese con su administrador", vbOKOnly + vbInformation
    Else
        tdbgDatos.Columns(14).Value = "1"
        cmdGrabar.Enabled = True
    End If
    
    If ColIndex = 15 Then 'COLUMNA NUMEWRO DE VOUCHER
        Cancel = 1
    End If
End Sub

Private Sub tdbgDatos_BeforeRowColChange(Cancel As Integer)
    With tdbgDatos
        If .Columns(14).Value = "1" Then
            If .Col = 11 And tdbgDatos.Columns(11) <> "__/__/____" Then
                ' *** Si fecha no esta completa, completarla
                tdbgDatos.Columns(11) = FormatoFecha(tdbgDatos.Columns(11))
                If VerificaFecha(tdbgDatos.Columns(11)) = False Then
                    .RefreshRow
                    .SetFocus
                    Cancel = 1
                End If
            End If
        End If
    End With
End Sub

Private Sub tdbgDatos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    If lArrDetalle(Bookmark, 16) = "1" Then
        RowStyle.BackColor = &HFF&
        RowStyle.ForeColor = &HFFFF&
    End If
    
    If lArrDetalle(Bookmark, 16) = "2" Then
        RowStyle.BackColor = gsColorDesactProv
    End If
    
End Sub

Private Sub tdbgDatos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 And (Mid(gsGrupo, 3, 1) = "1" Or gsGrupo = gsPrivilegioAdmin) Then
       tdbgDatos.Columns(11).Value = tdbgDatos.Columns(7).Value
       tdbgDatos.Columns(14).Value = "1"
    End If
    
    
End Sub

Private Sub Eliminar(Todo As Boolean)
    Dim strMensaje As String
    If IsNull(tdbgDatos.Bookmark) Then Exit Sub

    If lArrDetalle(tdbgDatos.Bookmark, 14) = "I" Then
        lArrDetalle.DeleteRows (tdbgDatos.Bookmark)
        tdbgDatos.ReBind
    Else
        If Todo = True Then
            strMensaje = "Deseas eliminar toda la conciliacion bancaria"
        Else
            strMensaje = "Deseas eliminar la conciliacion seleccionada"
        End If

        If MsgBox(strMensaje, vbYesNo + vbInformation, "Emininando Registro...") = vbYes Then
            Screen.MousePointer = vbHourglass
            If CierreMes(Me.tdbcMes.BoundText) = True Then
                Mensajes "El mes seleccionado ha sido cerrado. No se puede grabar.", vbInformation
                Exit Sub
            End If
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
    
            clsMante.InicializaClase
            clsMante.BeginTrans
        
            Call CargaArregloDet(tdbgDatos.Bookmark)
            lArrDet(0) = "ELIMINAR"
            
            If Todo = True Then
               lArrDet(0) = "ELIMINAR_ALL"
            End If
                
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaMovCheque", lArrDet(), False) = False Then
               Mensajes "El proceso no ha concluido. Verificar...", vbInformation
               clsMante.CancelTrans
               clsMante.FinalizaClase
               Exit Sub
            End If
            
            clsMante.CommitTrans
            clsMante.FinalizaClase
            Screen.MousePointer = vbNormal
            CargaTabla
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If ValidaCampos = False Then Exit Sub
    Eliminar (False)
End Sub

Private Sub tdbgDatosImp_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = 15 Then
        Cancel = 1
    End If
End Sub

Private Sub tdbgDatosImp_GotFocus()
    tdbgDatosImp.HighlightRowStyle = "HighlightRow"
End Sub
