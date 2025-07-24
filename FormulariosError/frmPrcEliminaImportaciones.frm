VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcEliminaImportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Importados"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   Icon            =   "frmPrcEliminaImportaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10050
   Begin VB.CheckBox chkSeleccion 
      Caption         =   "Seleccionar Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5355
      TabIndex        =   3
      Top             =   495
      Width           =   2085
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4305
      Left            =   240
      TabIndex        =   5
      Top             =   945
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   7594
      _LayoutType     =   4
      _RowHeight      =   20
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Interno"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Empresa"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Año"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Sucursal"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Periodo"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Libro"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Voucher"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fecha"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Glosa"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "InternoSuc"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   20
      Columns(10)._MaxComboItems=   5
      Columns(10).ValueItems(0)._DefaultItem=   0
      Columns(10).ValueItems(0).Value=   "1"
      Columns(10).ValueItems(0).Value.vt=   8
      Columns(10).ValueItems(0).DisplayValue=   "-1"
      Columns(10).ValueItems(0).DisplayValue.vt=   8
      Columns(10).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(10).ValueItems(1)._DefaultItem=   0
      Columns(10).ValueItems(1).Value=   "0"
      Columns(10).ValueItems(1).Value.vt=   8
      Columns(10).ValueItems(1).DisplayValue=   "0"
      Columns(10).ValueItems(1).DisplayValue.vt=   8
      Columns(10).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(10).ValueItems.Count=   2
      Columns(10).Caption=   "Sel"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2170"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1296"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1217"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=20"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1164"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1085"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=17"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=20"
      Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=1111"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1032"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=17"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=767"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=688"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=17"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1905"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1826"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=17"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=1984"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1905"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=17"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=7779"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=7699"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=20"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=423"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=344"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=20"
      Splits(0)._ColumnProps(53)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(55)=   "Column(10).Width=926"
      Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=847"
      Splits(0)._ColumnProps(58)=   "Column(10)._ColStyle=17"
      Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=35,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
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
      _StyleDefs(25)  =   "Splits(0).Style:id=79,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=88,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=80,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=81,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=82,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=84,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=83,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=86,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=87,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=89,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=90,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=94,.parent=79"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=80"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=81"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=83"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=98,.parent=79"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=80"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=81"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=83"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=102,.parent=79,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=80"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=81"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=83"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=106,.parent=79"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=80"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=81"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=83"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=110,.parent=79,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=80"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=81"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=83"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=114,.parent=79,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=80"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=81"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=83"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=118,.parent=79,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=115,.parent=80"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=116,.parent=81"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=117,.parent=83"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=122,.parent=79,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=119,.parent=80"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=120,.parent=81"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=121,.parent=83"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=126,.parent=79"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=123,.parent=80"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=124,.parent=81"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=125,.parent=83"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=130,.parent=79"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=127,.parent=80"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=128,.parent=81"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=129,.parent=83"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=134,.parent=79,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=131,.parent=80"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=132,.parent=81"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=133,.parent=83"
      _StyleDefs(81)  =   "Named:id=33:Normal"
      _StyleDefs(82)  =   ":id=33,.parent=0"
      _StyleDefs(83)  =   "Named:id=34:Heading"
      _StyleDefs(84)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   ":id=34,.wraptext=-1"
      _StyleDefs(86)  =   "Named:id=35:Footing"
      _StyleDefs(87)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   "Named:id=36:Selected"
      _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(90)  =   "Named:id=37:Caption"
      _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(92)  =   "Named:id=38:HighlightRow"
      _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(94)  =   "Named:id=39:EvenRow"
      _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(96)  =   "Named:id=40:OddRow"
      _StyleDefs(97)  =   ":id=40,.parent=33"
      _StyleDefs(98)  =   "Named:id=41:RecordSelector"
      _StyleDefs(99)  =   ":id=41,.parent=34"
      _StyleDefs(100) =   "Named:id=42:FilterBar"
      _StyleDefs(101) =   ":id=42,.parent=33"
   End
   Begin TDBDate6Ctl.TDBDate dtpDesde 
      Height          =   300
      Left            =   1125
      TabIndex        =   0
      Tag             =   "enabled"
      Top             =   180
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calendar        =   "frmPrcEliminaImportaciones.frx":0ECA
      Caption         =   "frmPrcEliminaImportaciones.frx":0FCC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcEliminaImportaciones.frx":1030
      Keys            =   "frmPrcEliminaImportaciones.frx":104E
      Spin            =   "frmPrcEliminaImportaciones.frx":10BA
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      Text            =   "03/08/2004"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   38202
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dtpHasta 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Tag             =   "enabled"
      Top             =   540
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calendar        =   "frmPrcEliminaImportaciones.frx":10E2
      Caption         =   "frmPrcEliminaImportaciones.frx":11E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcEliminaImportaciones.frx":1248
      Keys            =   "frmPrcEliminaImportaciones.frx":1266
      Spin            =   "frmPrcEliminaImportaciones.frx":12D2
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      Text            =   "03/08/2004"
      ValidateMode    =   0
      ValueVT         =   2010185735
      Value           =   38202
      CenturyMode     =   0
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   7830
      TabIndex        =   4
      Top             =   495
      Width           =   1680
      VariousPropertyBits=   268435483
      Caption         =   " Eliminar Asientos"
      PicturePosition =   327683
      Size            =   "2963;661"
      Picture         =   "frmPrcEliminaImportaciones.frx":12FA
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdMostrar 
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   495
      Width           =   1680
      Caption         =   " Mostrar"
      PicturePosition =   327683
      Size            =   "2963;635"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HASTA"
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
      Left            =   315
      TabIndex        =   7
      Top             =   585
      Width           =   570
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "DESDE"
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
      Left            =   315
      TabIndex        =   6
      Top             =   225
      Width           =   630
   End
End
Attribute VB_Name = "frmPrcEliminaImportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrDatos As New XArrayDB
Dim lArrCab() As Variant
Dim lItem As Integer
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkSeleccion_Click()
    Dim i As Integer
    
    If lItem <= 0 Then Exit Sub
    
    For i = 0 To arrDatos.Count(1) - 1
        'If i > 0 Then
            If chkSeleccion.Value = "1" Then
               arrDatos.Value(i, 10) = "1"
            Else
               arrDatos.Value(i, 10) = "0"
            End If
        'End If
    Next
    
    TDBGrid1.Update
    TDBGrid1.ReBind
    'TDBGrid1.Bookmark = TDBGrid1.Bookmark
    ' ***
End Sub

Private Sub cmdEliminaItem_Click()
    Dim i As Integer
    Dim respuesta As String
    Dim clsMante As clsMantoTablas
       
    If lItem <= 0 Then Exit Sub
    
    Me.TDBGrid1.Bookmark = Me.TDBGrid1.Bookmark
    
    ' *** Recorrer el arreglo y eliminar los seleccionados
    Me.TDBGrid1.Update
    respuesta = MsgBox("Desea eliminar los asientos seleccionados", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Asientos Seleccionados")
    If respuesta = vbNo Then Exit Sub
    ' ***
    
    Screen.MousePointer = vbHourglass
    For i = 0 To arrDatos.Count(1) - 1
        ' *** Aqui debo eliminar los asientos seleccionados
        If arrDatos(i, 10) = "1" Then
            ' *** Validar q mes no haya sido cerrado
            If CierreMes(arrDatos(i, 4)) = True Then
                Mensajes "Mes seleccionado ya fue cerrado. No se puede eliminar asiento.", vbInformation
                Exit For
            End If
            Set clsMante = New clsMantoTablas
            
            ' *** Revirtiendo Saldos
            Call CargaArregloAct(i)
            
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ActualizaSaldos", lArrCab(), False) = False Then
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            ' *** Eliminando el asiento
            Call CargaArregloCab(i)
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoCab", lArrCab(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Mensajes "Asientos han sido eliminados", vbInformation
    llenarGrilla
    ' ***
End Sub

Private Sub CargaArregloAct(Fila As Integer)
    ReDim lArrCab(7) As Variant
    lArrCab(0) = "REVERTIR"         ' *** Accion
    lArrCab(1) = arrDatos(Fila, 0)   ' *** Interno
    lArrCab(2) = arrDatos(Fila, 1)   ' *** Empresa
    lArrCab(3) = arrDatos(Fila, 2)   ' *** Año
    lArrCab(4) = arrDatos(Fila, 4)   ' *** Periodo
    lArrCab(5) = arrDatos(Fila, 5)   ' *** Libro
    lArrCab(6) = arrDatos(Fila, 6)   ' *** Voucher
    lArrCab(7) = gsUsuario
End Sub

Private Sub CargaArregloCab(Fila As Integer)
    ReDim lArrCab(15) As Variant
    lArrCab(0) = "ELIMINAR"             ' Accion
    lArrCab(1) = arrDatos(Fila, 0)   ' Numero Interno
    lArrCab(2) = arrDatos(Fila, 1)   ' Empresa
    lArrCab(3) = arrDatos(Fila, 2)   ' Año
    lArrCab(4) = arrDatos(Fila, 4)   ' Periodo
    lArrCab(5) = arrDatos(Fila, 5)   ' Libro
    lArrCab(6) = arrDatos(Fila, 6)   ' Voucher
    lArrCab(7) = "01/01/1900"
    lArrCab(8) = ""
    lArrCab(9) = ""
    lArrCab(10) = 0
    lArrCab(11) = ""
    lArrCab(12) = ""
    lArrCab(13) = ""
    lArrCab(14) = ""
    lArrCab(15) = ""
End Sub

Private Sub cmdMostrar_Click()
    llenarGrilla
    chkSeleccion.Value = "0"
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
        
    Call Centrar_form(Me)
    
    dtpDesde = FechaServidor
    dtpHasta = dtpDesde
    llenarGrilla
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdEliminaItem.Enabled = False
        
    Else
        Me.cmdEliminaItem.Enabled = True
        
    End If
    
End Sub

Private Sub llenarGrilla()
    ' *** Aqui definir la tabla y los campos
    Dim psql$
    arrDatos.Clear
    
    psql$ = " SELECT Ase_cNummov, Emp_cCodigo, Pan_cAnio, Ase_cCodSucursal, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, " & _
            " Ase_dFecha, Ase_cGlosa, Ase_cNumMovTra, '0' as tipo FROM dbo.CNC_ASIENTO_VOUCHER " & _
            " WHERE Ase_cDeleted <> '*' AND Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
            " AND (CONVERT(DATETIME, Ase_dFecha, 103) >= CONVERT(DATETIME, '" & dtpDesde & "', 103)" & _
            " AND CONVERT(DATETIME, Ase_dFecha, 103) <= CONVERT(DATETIME, '" & dtpHasta & "', 103))" & _
            " AND Ase_cCodSoft <> '' ORDER BY Ase_cCodSucursal, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher  "
            
    Call LlenarArregloRetornandoFilas(arrDatos, psql$, lItem)
    Set Me.TDBGrid1.Array = arrDatos
    
    Me.TDBGrid1.ReBind
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        'Call SeteaFondoForm(Me)
        
        With TDBGrid1
            .Width = Me.Width - 500
            .Height = Me.Height - .Top - 500
        End With

    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set arrDatos = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

