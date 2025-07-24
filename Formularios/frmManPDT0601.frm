VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManPDT0601 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDT 0601 - Exportación de Datos"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17445
   Icon            =   "frmManPDT0601.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   17445
   Begin TrueOleDBGrid70.TDBGrid tdbgPDT0601 
      Height          =   4860
      Left            =   90
      TabIndex        =   0
      Top             =   945
      Width           =   17205
      _ExtentX        =   30348
      _ExtentY        =   8573
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Número de R.U.C."
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Entidad"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cod. Tipo de Documento"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo de Documento"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Serie Documento"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Número Documento"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fecha Emisión"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "External Editor"
      Columns(6).ExternalEditor=   "tdbdFecha"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fecha Pago"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Importe"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Cod. Tipo Comp. Emi."
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Importe AFP/ONP"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Tipo Comprobante Emitido"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Sel."
      Columns(12).DataField=   ""
      Columns(12).DataWidth=   1
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   20
      Columns(13)._MaxComboItems=   5
      Columns(13).ValueItems(0)._DefaultItem=   0
      Columns(13).ValueItems(0).Value=   "1"
      Columns(13).ValueItems(0).Value.vt=   8
      Columns(13).ValueItems(0).DisplayValue=   "-1"
      Columns(13).ValueItems(0).DisplayValue.vt=   8
      Columns(13).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(13).ValueItems(1)._DefaultItem=   0
      Columns(13).ValueItems(1).Value=   "0"
      Columns(13).ValueItems(1).Value.vt=   8
      Columns(13).ValueItems(1).DisplayValue=   "0"
      Columns(13).ValueItems(1).DisplayValue.vt=   8
      Columns(13).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(13).ValueItems.Count=   2
      Columns(13).Caption=   "Retención 4ta"
      Columns(13).DataField=   ""
      Columns(13).DropDown=   "tdbdDetracciones"
      Columns(13).DropDown.vt=   8
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8705"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6350"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6271"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8704"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8705"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1693"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1614"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8705"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1931"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1852"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2170"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2090"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2328"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2249"
      Splits(0)._ColumnProps(39)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=8705"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=2143"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=2064"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=8706"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(52)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(10)._ColStyle=8708"
      Splits(0)._ColumnProps(56)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(58)=   "Column(11).Width=2619"
      Splits(0)._ColumnProps(59)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(11)._WidthInPix=2540"
      Splits(0)._ColumnProps(61)=   "Column(11)._ColStyle=513"
      Splits(0)._ColumnProps(62)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(63)=   "Column(12).Width=900"
      Splits(0)._ColumnProps(64)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(12)._WidthInPix=820"
      Splits(0)._ColumnProps(66)=   "Column(12)._ColStyle=513"
      Splits(0)._ColumnProps(67)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(68)=   "Column(13).Width=1640"
      Splits(0)._ColumnProps(69)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(13)._WidthInPix=1561"
      Splits(0)._ColumnProps(71)=   "Column(13)._ColStyle=513"
      Splits(0)._ColumnProps(72)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(73)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(77)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(78)=   "Column(14).Order=15"
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
      CellTips        =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=162,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
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
      _StyleDefs(25)  =   "Splits(0).Style:id=25,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=60,.parent=4,.namedParent=37"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=30,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=29,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=31,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=32,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=59,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=61,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=62,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=66,.parent=25,.alignment=2,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=26"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=27"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=29"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=70,.parent=25,.alignment=0,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=26"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=27"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=29"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=16,.parent=25,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=26"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=27"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=29"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=74,.parent=25,.alignment=2,.bgcolor=&HFFFFFF&"
      _StyleDefs(50)  =   ":id=74,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=26"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=27"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=29"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=78,.parent=25,.alignment=2,.bgcolor=&HFFFFFF&"
      _StyleDefs(55)  =   ":id=78,.locked=-1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=26"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=27"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=29"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=126,.parent=25,.alignment=2,.bgcolor=&HFFFFFF&"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=26"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=27"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=29"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=130,.parent=25,.alignment=2,.bgcolor=&HFFFFFF&"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=26"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=27"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=29"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=134,.parent=25,.alignment=2,.locked=-1"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=26"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=27"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=29"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=138,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(72)  =   ":id=138,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=135,.parent=26"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=136,.parent=27"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=137,.parent=29"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=46,.parent=25"
      _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=26"
      _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=27"
      _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=29"
      _StyleDefs(80)  =   "Splits(0).Columns(10).Style:id=20,.parent=25,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=26"
      _StyleDefs(82)  =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=27"
      _StyleDefs(83)  =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=29"
      _StyleDefs(84)  =   "Splits(0).Columns(11).Style:id=58,.parent=25,.alignment=2"
      _StyleDefs(85)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=26"
      _StyleDefs(86)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=27"
      _StyleDefs(87)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=29"
      _StyleDefs(88)  =   "Splits(0).Columns(12).Style:id=24,.parent=25,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(89)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(90)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(91)  =   "Splits(0).Columns(12).HeadingStyle:id=21,.parent=26"
      _StyleDefs(92)  =   "Splits(0).Columns(12).FooterStyle:id=22,.parent=27"
      _StyleDefs(93)  =   "Splits(0).Columns(12).EditorStyle:id=23,.parent=29"
      _StyleDefs(94)  =   "Splits(0).Columns(13).Style:id=142,.parent=25,.alignment=2"
      _StyleDefs(95)  =   "Splits(0).Columns(13).HeadingStyle:id=139,.parent=26"
      _StyleDefs(96)  =   "Splits(0).Columns(13).FooterStyle:id=140,.parent=27"
      _StyleDefs(97)  =   "Splits(0).Columns(13).EditorStyle:id=141,.parent=29"
      _StyleDefs(98)  =   "Splits(0).Columns(14).Style:id=50,.parent=25"
      _StyleDefs(99)  =   "Splits(0).Columns(14).HeadingStyle:id=47,.parent=26"
      _StyleDefs(100) =   "Splits(0).Columns(14).FooterStyle:id=48,.parent=27"
      _StyleDefs(101) =   "Splits(0).Columns(14).EditorStyle:id=49,.parent=29"
      _StyleDefs(102) =   "Named:id=33:Normal"
      _StyleDefs(103) =   ":id=33,.parent=0"
      _StyleDefs(104) =   "Named:id=34:Heading"
      _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(106) =   ":id=34,.wraptext=-1"
      _StyleDefs(107) =   "Named:id=35:Footing"
      _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(109) =   "Named:id=36:Selected"
      _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(111) =   "Named:id=37:Caption"
      _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(113) =   "Named:id=38:HighlightRow"
      _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=39:EvenRow"
      _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(117) =   "Named:id=40:OddRow"
      _StyleDefs(118) =   ":id=40,.parent=33"
      _StyleDefs(119) =   "Named:id=41:RecordSelector"
      _StyleDefs(120) =   ":id=41,.parent=34,.alignment=3"
      _StyleDefs(121) =   "Named:id=42:FilterBar"
      _StyleDefs(122) =   ":id=42,.parent=33,.alignment=3"
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   135
      Width           =   3435
      _ExtentX        =   6059
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
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=688"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=609"
      Splits(0)._ColumnProps(11)=   "Column(1)._VertColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
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
      _PropDict       =   $"frmManPDT0601.frx":0ECA
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
   Begin TrueOleDBList70.TDBCombo tdbcLibro 
      Height          =   300
      Left            =   855
      TabIndex        =   2
      Tag             =   "enabled"
      Top             =   495
      Width           =   3435
      _ExtentX        =   6059
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
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AllowRowSizing=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2196"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2117"
      Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2196"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2117"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=2196"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2117"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2196"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2117"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=2196"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=2117"
      Splits(0)._ColumnProps(56)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
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
      _PropDict       =   $"frmManPDT0601.frx":0F51
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
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(72)  =   "Named:id=33:Normal"
      _StyleDefs(73)  =   ":id=33,.parent=0"
      _StyleDefs(74)  =   "Named:id=34:Heading"
      _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   ":id=34,.wraptext=-1"
      _StyleDefs(77)  =   "Named:id=35:Footing"
      _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=36:Selected"
      _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=37:Caption"
      _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(83)  =   "Named:id=38:HighlightRow"
      _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=39:EvenRow"
      _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(87)  =   "Named:id=40:OddRow"
      _StyleDefs(88)  =   ":id=40,.parent=33"
      _StyleDefs(89)  =   "Named:id=41:RecordSelector"
      _StyleDefs(90)  =   ":id=41,.parent=34"
      _StyleDefs(91)  =   "Named:id=42:FilterBar"
      _StyleDefs(92)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Digíte una ""x"" en la columna ""Sel."" para seleccionar los documentos a Exportar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   4320
      TabIndex        =   13
      Top             =   240
      Width           =   4365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
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
      Left            =   135
      TabIndex        =   12
      Top             =   180
      Width           =   645
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
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   540
      Width           =   420
   End
   Begin MSForms.CommandButton cmdListar 
      Height          =   375
      Left            =   12615
      TabIndex        =   10
      ToolTipText     =   "Cargar nueva Configuración"
      Top             =   105
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":0FD8
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   8655
      TabIndex        =   9
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   45
      Visible         =   0   'False
      Width           =   1575
      Caption         =   " Insertar Item"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":1572
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      ToolTipText     =   "Eliminar el movimientos seleccionado"
      Top             =   45
      Visible         =   0   'False
      Width           =   1575
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2778;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
      Top             =   450
      Visible         =   0   'False
      Width           =   1575
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":1B0C
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   12615
      TabIndex        =   6
      ToolTipText     =   "Grabar modificaciones"
      Top             =   510
      Width           =   1575
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":20A6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdTodos 
      Height          =   375
      Left            =   8655
      TabIndex        =   5
      ToolTipText     =   "Insertar todos los movimientos del libro y mes seleccionado"
      Top             =   450
      Visible         =   0   'False
      Width           =   1575
      Caption         =   " Insertar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":2640
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   14265
      TabIndex        =   4
      Top             =   90
      Width           =   1575
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":2BDA
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPreliminar 
      Height          =   375
      Left            =   14265
      TabIndex        =   3
      Top             =   495
      Width           =   1575
      Caption         =   " Vista Preliminar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDT0601.frx":2F74
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmManPDT0601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Dim lArrDatos As New XArrayDB

Dim NUM_FILAS As Integer
Dim NUM_COLUMNAS As Integer

Dim nCol_cRuc As Integer
Dim nCol_cEntidad As Integer
Dim nCol_cTipDocu As Integer
Dim nCol_cDscTipoDocu As Integer
Dim nCol_cSerDocu As Integer
Dim nCol_cNroDoc As Integer
Dim nCol_cFechaEmision As Integer
Dim nCol_cFechaPago As Integer
Dim nCol_cImporte As Integer
Dim nCol_cCodTipoCompEmi As Integer
Dim nCol_cTipoCompEmi As Integer
Dim nCol_cRetencion As Integer
Dim lArrDet() As Variant
Dim rsArreglo As ADODB.Recordset
Dim CorItem As Integer
Public InsItem As Boolean
Dim DelAll As Boolean
Dim lControl As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property
Private Sub cmdEliminaItem_Click()
 If ValidaCampos = False Then Exit Sub

 If CE(tdbgPDT0601.Columns(1)) <> "" Then
    cmdEliminaItem.Enabled = False
    DoEvents
    
    EliminarItem
    lArrDatos.DeleteRows (tdbgPDT0601.Bookmark)

    Set tdbgPDT0601.Array = lArrDatos
    
    Call UpdateGrilla(tdbgPDT0601)
    Call RebindGrilla(tdbgPDT0601)
    
    DoEvents
    cmdEliminaItem.Enabled = True
    
 End If

 Call UpdateGrilla(tdbgPDT0601)
 Call RefreshGrilla(tdbgPDT0601)
End Sub

Private Sub cmdEliminarTodo_Click()
On Error GoTo Control

 DoEvents
  Call EliminaTodo
 DoEvents

 Set rsArreglo = Nothing
 lArrDatos.Clear
 lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
        
 Set tdbgPDT0601.Array = lArrDatos

 Call UpdateGrilla(tdbgPDT0601)
 Call RebindGrilla(tdbgPDT0601)
  
 DelAll = False
 
Exit Sub
Control:
 MsgBox Err.Description
End Sub

Private Sub cmdGrabar_Click()
If tdbgPDT0601.ApproxCount = 0 Then Exit Sub
 If ValidaCampos = False Then Exit Sub
 If ValidaSeleccionCampos = False Then MsgBox "Debe seleccionar por lo menos una fila para Grabar", vbInformation, App.Title: Exit Sub
 cmdGrabar.Enabled = False
 DoEvents
 If Grabar = True Then
  Mensajes "Datos se grabaron con exito.", vbInformation
 End If
 DoEvents
 cmdGrabar.Enabled = True
End Sub
Private Sub cmdInsertarItem_Click()
On Error GoTo Control

 If ValidaCampos = False Then Exit Sub
 
If Not InsItem Then

 Set rsArreglo = Nothing
 lArrDatos.Clear
 lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas

 Set tdbgPDT0601.Array = lArrDatos

 Call UpdateGrilla(tdbgPDT0601)
 Call RebindGrilla(tdbgPDT0601)
 
End If

 cmdInsertarItem.Enabled = False
 DoEvents
  frmBuscaDocsPDT0601.Show
 DoEvents
 cmdInsertarItem.Enabled = True
 
Exit Sub
Control:
 MsgBox Err.Description
End Sub

Private Sub cmdListar_Click()
 If ValidaCampos = False Then Exit Sub
 InsItem = False
 cmdListar.Enabled = False
 DoEvents
 If DelAll Then DelAll = False
 Call llenaGrilla
 DoEvents
 cmdListar.Enabled = True
' Dim i As Integer
' For i = 1 To tdbgPDT0601.Columns.Count
'    tdbgPDT0601.Columns(11).Value = ""
' Next i
End Sub

Private Sub cmdPreliminar_Click()
  cmdPreliminar.Enabled = False
  DoEvents
  
  Dim matriz(8) As Variant
  Dim Titulo As String
  Titulo = "Reporte Exportación de Datos PDT0601"
  Titulo = UCase(Titulo)
  matriz(0) = "@Accion;VISTA_PRELIMINAR;True"
  matriz(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
  matriz(2) = "@Pan_cAnio;" & gsAnio & ";True"
  matriz(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
  matriz(4) = "@Lib_cTipoLibro;" & IIf(Trim(tdbcLibro.BoundText) = "00", "", tdbcLibro.BoundText) & ";True"
  
  matriz(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
  matriz(6) = "@RUC;" & "RUC : " & gsRUC & ";True"
  matriz(7) = "@NOMBREMES;" & NombreMes(tdbcMes.BoundText) & ";True"
  
  Dim formulas(0) As Variant
  AbreReporteParam gsDSN, Me, rutaReportes & "RptExpPDT0601.rpt", crptToWindow, Titulo, "", matriz(), formulas()

  cmdPreliminar.Enabled = True
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub cmdTodos_Click()
DelAll = True

 If ValidaCampos = False Then Exit Sub
    
 Set rsArreglo = Nothing
 lArrDatos.Clear
 lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
        
 Set tdbgPDT0601.Array = lArrDatos
        
 Call UpdateGrilla(tdbgPDT0601)
 Call RebindGrilla(tdbgPDT0601)
    
 cmdTodos.Enabled = False
 DoEvents
   Call llenaGrilla
 DoEvents
 cmdTodos.Enabled = True
DelAll = False
End Sub

Private Sub Form_Load()
On Error GoTo Control

 Dim sqlcombos As String
    
 Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
 Call IniciaVariables
 
 NUM_FILAS = 0
 NUM_COLUMNAS = 15
 
 pCargaCfgLibro
 
 Call Centrar_form(Me)
 Call LlenaComboMesAddItem(tdbcMes, True, True, "[ Seleccione Mes]")
    
 Dim registros As Integer
 
              'and LIB_CTIPOLIBRO='" & lsLibroCom & "' " &
              
 'sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
 '            "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
 '            "AND LIB_CTIPOLIBRO IN('06','03') " & _
 '            "ORDER BY LIB_CDESCRIPCION "
    
' registros = LlenarComboAddItem(tdbcLibro, sqlcombos, True, True, "[ Seleccione Libro ]")
  tdbcLibro.AddItem "00" + ";" + "[TODOS LOS LIBROS]"
  
  tdbcLibro.Bookmark = 0
  tdbcLibro.ListField = "column1"
  tdbcLibro.BoundColumn = "column0"
  tdbcLibro.ReBind
  
'If registros > 0 Then
    DoEvents
    
    Call llenaGrilla

    tdbcLibro.Enabled = True

'Else
'    Mensajes "No se crearon los Libros Contables en el Sistema, Ingreselos en Mantenimiento de Libros", vbOKOnly + vbInformation
'    DesactivaBotones False
'End If

Call ConfigurarColumnas

lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
InsItem = False
    
Exit Sub
Control:
 MsgBox Err.Description
End Sub
Sub IniciaVariables()
 nCol_cRuc = 0
 nCol_cEntidad = 1
 nCol_cTipDocu = 2
 nCol_cDscTipoDocu = 3
 nCol_cSerDocu = 4
 nCol_cNroDoc = 5
 nCol_cFechaEmision = 6
 nCol_cFechaPago = 7
 nCol_cImporte = 8
 nCol_cCodTipoCompEmi = 9
 nCol_cTipoCompEmi = 10
 nCol_cRetencion = 11
End Sub
Private Sub DesactivaBotones(Valor As Boolean)
 Me.cmdEliminaItem.Enabled = Valor
 Me.cmdEliminarTodo.Enabled = Valor
 Me.cmdGrabar.Enabled = Valor
 Me.cmdInsertarItem.Enabled = Valor
 Me.cmdTodos.Enabled = Valor
    
 DoEvents
End Sub
Private Sub ConfigurarColumnas()
 Call OcultaColumna(nCol_cTipDocu)
 Call OcultaColumna(nCol_cCodTipoCompEmi)
End Sub
Private Sub OcultaColumna(nCol As Integer)
 On Error GoTo serror
    With tdbgPDT0601
        .Columns(nCol).Visible = False
        .Columns(nCol).Width = 0
        .Splits(0).Columns(nCol).AllowFocus = False
        .Splits(0).Columns(nCol).AllowSizing = False
    End With
    Exit Sub
serror:
 MsgBox Err.Description
End Sub

Private Sub BloqueaColumna(nCol As Integer)
    On Error GoTo serror
    With tdbgPDT0601
        .Columns(nCol).BackColor = gsColorDesactivado
        .Splits(0).Columns(nCol).AllowFocus = False
        .Splits(0).Columns(nCol).AllowSizing = True
    End With
    Exit Sub
serror:
End Sub
Public Sub llenaGrilla()
On Error GoTo Control
    
 Dim sqlcombos As String
 Dim clDatos As clsMantoTablas
 
Screen.MousePointer = vbHourglass

 Set clDatos = New clsMantoTablas
 
 Dim arrDatos() As Variant
 Dim i As Integer
 Dim Col As Integer
    
 If tdbcMes.BoundText = "" Then
  pSetFocus tdbcMes
  Screen.MousePointer = vbDefault
  Exit Sub
 End If
        
 If tdbcLibro.BoundText = "" Then
  pSetFocus tdbcLibro
  Screen.MousePointer = vbDefault
  Exit Sub
 End If
    
 If Not DelAll Then
  sqlcombos = "spCn_GrabaPDT0601 'BUSCARTODOS','" & gsEmpresa & "','" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', ''"
 Else
  sqlcombos = "spCn_GrabaPDT0601 'SELECCIONA_TODOS','" & gsEmpresa & "','" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', ''"
 End If

 Set rsArreglo = New ADODB.Recordset
 
 arrDatos = Array(sqlcombos)
 Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
 
 If rsArreglo Is Nothing Then
  Screen.MousePointer = vbNormal
  Set rsArreglo = Nothing
  lArrDatos.Clear
  lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
        
  Set tdbgPDT0601.Array = lArrDatos
        
  Call UpdateGrilla(tdbgPDT0601)
  Call RebindGrilla(tdbgPDT0601)

  Exit Sub
 End If
    
 lArrDatos.Clear
    
 lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
    
 If lArrDatos.Count(1) = 0 And lArrDatos.Count(2) = 0 Then
  lArrDatos.AppendRows
        
  Call UpdateGrilla(tdbgPDT0601)
  Call RebindGrilla(tdbgPDT0601)
        
 End If
    
 i = 0
 Col = -1
 Do While Not rsArreglo.EOF
            
  lArrDatos(i, 1 + Col) = CE(rsArreglo(0).Value)
  lArrDatos(i, 2 + Col) = CE(rsArreglo(1).Value)
  lArrDatos(i, 3 + Col) = CE(rsArreglo(2).Value)
  lArrDatos(i, 4 + Col) = CE(rsArreglo(3).Value)
  lArrDatos(i, 5 + Col) = CE(rsArreglo(4).Value)
  lArrDatos(i, 6 + Col) = CE(rsArreglo(5).Value)
  lArrDatos(i, 7 + Col) = CE(rsArreglo(6).Value)
  lArrDatos(i, 8 + Col) = CE(rsArreglo(7).Value)
  lArrDatos(i, 9 + Col) = CE(rsArreglo(8).Value)
  lArrDatos(i, 10 + Col) = CE(rsArreglo(9).Value)
  lArrDatos(i, 11 + Col) = CE(rsArreglo(10).Value)
  lArrDatos(i, 12 + Col) = CE(rsArreglo(11).Value)
  lArrDatos(i, 13 + Col) = CE(rsArreglo(12).Value)
  lArrDatos(i, 14 + Col) = CE(rsArreglo(13).Value)
  lArrDatos(i, 15 + Col) = CE(rsArreglo(14).Value)

  lArrDatos.AppendRows
        
  rsArreglo.MoveNext
        
  i = i + 1
 Loop
    
 Set tdbgPDT0601.Array = lArrDatos
    
 Call UpdateGrilla(tdbgPDT0601)
 Call RebindGrilla(tdbgPDT0601)
    
Screen.MousePointer = vbDefault

Exit Sub

Control:
 MsgBox Err.Description
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        tdbgPDT0601.Width = Me.Width - 200
        tdbgPDT0601.Height = Me.Height - 1600
    End If
    
    Exit Sub
serror:
 Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set lArrDatos = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub tdbcLibro_ItemChange()
 Call llenaGrilla
End Sub

Private Sub tdbcMes_ItemChange()
 Call llenaGrilla
End Sub
Private Function ValidaCampos() As Boolean
 ValidaCampos = False

 If tdbcMes.BoundText = "" Then
  Mensajes "Seleccione un Periodo a Consultar"
  pSetFocus tdbcMes
  Exit Function
 End If
    
 If tdbcLibro.BoundText = "" Then
  Mensajes "Seleccione un  Libro a Consultar"
  pSetFocus tdbcLibro
  Exit Function
 End If

 ValidaCampos = True
End Function
Private Function ValidaSeleccionCampos() As Boolean
'PGBV 28102013
On Error GoTo Control

 Dim i As Integer
 tdbgPDT0601.Row = 0
 
 If Trim(tdbgPDT0601.Columns(0).Value) = "" Then Exit Function
 ValidaSeleccionCampos = False
 
 If Not InsItem Then
  rsArreglo.MoveFirst
  For i = 0 To rsArreglo.RecordCount - 1
   If Trim(lArrDatos(i, 13 + -1)) = "x" Then
    ValidaSeleccionCampos = True
   End If
   rsArreglo.MoveNext
  Next i
 Else
  For i = 0 To tdbgPDT0601.ApproxCount - 1
   If Trim(lArrDatos(i, 13 + -1)) = "x" Then
    ValidaSeleccionCampos = True
   End If
  Next i
 End If
Exit Function

Control:
 MsgBox Err.Description
End Function
Public Function Grabar() As Boolean
On Error GoTo ERROR

 Dim i As Integer
 Dim clsMante As New clsMantoTablas
        
 Grabar = True
 Dim Fila As Integer
    
 Fila = CuentaFilas
 '--------------------------------------------
 If Not InsItem Then
  DoEvents
   Call EliminaTodo
  DoEvents
 End If
 '--------------------------------------------
 Screen.MousePointer = vbHourglass
 
 If Not InsItem Then
     clsMante.InicializaClase
     clsMante.BeginTrans
    
      For i = 0 To Fila - 1
        If CE(lArrDatos(i, 12)) = "x" Then
          Call CargaArregloDet(i)
         If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPDT0601", lArrDet(), False) = False Then
           Mensajes "El proceso no ha concluido. Verificar fila..." & i, vbInformation
           Screen.MousePointer = vbDefault
                        
           clsMante.CancelTrans
           clsMante.FinalizaClase
           Set clsMante = Nothing
           Grabar = False
                        
          Exit Function
         End If
        End If
      Next
        
     clsMante.CommitTrans
     clsMante.FinalizaClase
 Else
  For i = 0 To Fila - 1
   If CE(lArrDatos(i, 12)) = "x" Then
    GrabarItem
    tdbgPDT0601.Row = tdbgPDT0601.Row + 1
   End If
  Next i
  InsItem = False
 End If
 Set clsMante = Nothing
    
 CorItem = 0
 Screen.MousePointer = vbDefault
 Exit Function

ERROR:
 Grabar = False
 Screen.MousePointer = vbNormal
End Function
Private Function CuentaFilas() As Integer

 Dim i As Integer
 Dim Contador As Integer
 Contador = 0

 For i = 0 To lArrDatos.Count(1) - 1
  If CE(lArrDatos(i, 1)) <> "" Then
   Contador = Contador + 1
  End If
 Next i
 
 CuentaFilas = Contador  'lArrDatos.UpperBound(1) - lArrDatos.LowerBound(1)
 
End Function
Private Sub EliminaTodo()
On Error GoTo Control
 Dim clsMante As New clsMantoTablas
        
 Call EliminaArreglo
        
 clsMante.InicializaClase
 clsMante.BeginTrans
        
 If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPDT0601", lArrDet(), False) = False Then
  Mensajes "El proceso no ha concluido....", vbInformation
  Screen.MousePointer = vbDefault
            
  clsMante.CancelTrans
  clsMante.FinalizaClase
            
  cmdEliminarTodo.Enabled = True
            
  Exit Sub
 End If
        
 clsMante.CommitTrans
 clsMante.FinalizaClase
 Set clsMante = Nothing
 
Exit Sub

Control:
 MsgBox Err.Description
 Screen.MousePointer = vbDefault
End Sub
Private Sub EliminaArreglo()
 ReDim lArrDet(5)
 lArrDet(0) = "ELIMINAR_TODOS"
 lArrDet(1) = gsEmpresa
 lArrDet(2) = gsAnio
 lArrDet(3) = tdbcMes.BoundText
 lArrDet(4) = tdbcLibro.BoundText
End Sub
Private Sub CargaArregloDet(item As Integer)
On Error GoTo Control

Dim i As Integer
i = 0
CorItem = CorItem + 1
 ReDim lArrDet(18) As Variant
 
  lArrDet(0) = "INSERTAR"
  lArrDet(1) = gsEmpresa                    'Empresa
  lArrDet(2) = gsAnio                       'Año
  lArrDet(3) = tdbcMes.BoundText            'Periodo
  lArrDet(4) = tdbcLibro.BoundText          'Libro
  lArrDet(5) = CorItem                      'Item
  lArrDet(6) = CE(lArrDatos(item, 2 + i))   'Tipo Documento Entidad
  lArrDet(7) = CE(lArrDatos(item, 0 + i))   'RUC
'TATA-004
  'lArrDet(8) = CE(lArrDatos(item, 9 + i))   'Tipo de Documento Emitido
  lArrDet(8) = CE(lArrDatos(item, 11 + i))   'Tipo de Documento Emitido
'FIN-TATA-004
  lArrDet(9) = CE(lArrDatos(item, 4 + i))   'Serie
  lArrDet(10) = CE(lArrDatos(item, 5 + i))  'Numero
  lArrDet(11) = CE(lArrDatos(item, 8 + i))  'Importe
  lArrDet(12) = CE(lArrDatos(item, 6 + i))  'Fecha Emision
  lArrDet(13) = CE(lArrDatos(item, 7 + i))  'Fecha Pago
  lArrDet(14) = CE(lArrDatos(item, 13 + i)) 'Retencion
  lArrDet(15) = gsUsuario                   'Usuario
  lArrDet(16) = CE(lArrDatos(item, 10 + i))  'Importe AFP/ONP
  lArrDet(17) = CE(lArrDatos(item, 14 + i))  'Tipo de Retencion AFP-ONP-NA
Exit Sub
Control:
 MsgBox Err.Description
End Sub

Sub LlenaDatos()
On Error GoTo Control

 Dim i As Integer
 Dim Fila As Integer

    i = 0
    Fila = CuentaFilas

    If lArrDatos.Count(1) = Fila + 1 Or lArrDatos.Count(1) = Fila Then
        Call AgregaFila
    End If

    With frmBuscaDocsPDT0601.tdbgPDT0601
     If Trim(tdbgPDT0601.Columns(0).Value) = Trim(CE(.Columns(0).Value)) And Trim(tdbgPDT0601.Columns(5).Value) = Trim(CE(.Columns(5).Value)) Then
      MsgBox "Este Item ya se encuentra agregado, verificar...", vbInformation, App.Title
      InsItem = False
      Exit Sub
     End If
        lArrDatos(Fila, 0 + i) = CE(.Columns(0).Value)
        lArrDatos(Fila, 1 + i) = CE(.Columns(1).Value)
        lArrDatos(Fila, 2 + i) = CE(.Columns(2).Value)
        lArrDatos(Fila, 3 + i) = CE(.Columns(3).Value)
        lArrDatos(Fila, 4 + i) = CE(.Columns(4).Value)
        lArrDatos(Fila, 5 + i) = CE(.Columns(5).Value)
        lArrDatos(Fila, 6 + i) = CE(.Columns(6).Value)
        lArrDatos(Fila, 7 + i) = CE(.Columns(7).Value)
        lArrDatos(Fila, 8 + i) = CE(.Columns(8).Value)
        lArrDatos(Fila, 9 + i) = CE(.Columns(9).Value)
        lArrDatos(Fila, 10 + i) = CE(.Columns(10).Value)
        lArrDatos(Fila, 11 + i) = "x"
    End With

    Set tdbgPDT0601.Array = lArrDatos
    'GrabarItem
    Call UpdateGrilla(tdbgPDT0601)
    Call RebindGrilla(tdbgPDT0601)

    Call AgregaFila
    tdbgPDT0601.Row = tdbgPDT0601.ApproxCount - 2
    InsItem = False
    Call frmBuscaDocsPDT0601.Form_Unload(0)

Exit Sub
Control:
 MsgBox Err.Description
End Sub
Private Sub AgregaFila()
 Dim Filas As Integer
 Filas = CuentaFilas
 lArrDatos.ReDim 0, Filas, 0, NUM_COLUMNAS    ' filas
    
 Call UpdateGrilla(tdbgPDT0601)
 Call RebindGrilla(tdbgPDT0601)
End Sub
Sub GrabarItem()


Dim clsMante As New clsMantoTablas
Dim RstNroReg As ADODB.Recordset

Dim arrDatos() As Variant
Dim ItemReg As Integer
Dim sql As String
Dim i As Integer

Set RstNroReg = New ADODB.Recordset

sql = "spCn_NroRegPDT0601 '" & gsEmpresa & "','" & gsAnio & "', '" & tdbcMes.BoundText & "', '" & tdbcLibro.BoundText & "'"
   
 arrDatos = Array(sql)
 Set RstNroReg = clsMante.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
 
If Not RstNroReg Is Nothing Then
 ItemReg = RstNroReg.Fields(0) + 1
End If

Set clsMante = New clsMantoTablas

i = 0

 ReDim lArrDet(18) As Variant
 
With tdbgPDT0601
  lArrDet(0) = "INSERTAR"
  lArrDet(1) = gsEmpresa                                    'Empresa
  lArrDet(2) = gsAnio                                       'Año
  lArrDet(3) = tdbcMes.BoundText                            'Periodo
  lArrDet(4) = tdbcLibro.BoundText                          'Libro
  lArrDet(5) = ItemReg                                      'Item
  lArrDet(6) = CE(.Columns(2).Value)                        'Tipo Documento Entidad
  lArrDet(7) = CE(.Columns(0).Value)                        'RUC
  lArrDet(8) = CE(.Columns(9).Value)                        'Tipo de Documento Emitido
  lArrDet(9) = CE(.Columns(4).Value)                        'Serie
  lArrDet(10) = CE(.Columns(5).Value)                       'Numero
  lArrDet(11) = CE(.Columns(8).Value)                       'Importe
  lArrDet(12) = CE(.Columns(6).Value)                       'Fecha Emision
  lArrDet(13) = CE(.Columns(7).Value)                       'Fecha Pago
  lArrDet(14) = IIf(CE(.Columns(13).Value) = "", "0", "1") 'Retencion
  lArrDet(15) = gsUsuario                                   'Usuario
  lArrDet(16) = CE(.Columns(14).Value)  'Importe AFP/ONP
  lArrDet(17) = IIf(CE(.Columns(15).Value) = 3, "", CE(.Columns(15).Value))  'Tipo de Retencion AFP-ONP-NA
End With
 clsMante.InicializaClase
 clsMante.BeginTrans
  
  If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPDT0601", lArrDet(), False) = False Then
   Mensajes "El proceso no ha concluido. Verificar fila..." & i, vbInformation
                
   clsMante.CancelTrans
   clsMante.FinalizaClase
   Set clsMante = Nothing
                
   Exit Sub
  End If

 clsMante.CommitTrans
 clsMante.FinalizaClase

 Set clsMante = Nothing
End Sub
Sub EliminarItem()
On Error GoTo Control
 Dim clsMante As New clsMantoTablas
        
 ReDim lArrDet(11)
 lArrDet(0) = "ELIMINAR_ITEM"
 lArrDet(1) = gsEmpresa
 lArrDet(2) = gsAnio
 lArrDet(3) = tdbcMes.BoundText
 lArrDet(4) = tdbcLibro.BoundText
 lArrDet(7) = CE(tdbgPDT0601.Columns(0).Value)
 lArrDet(8) = CE(tdbgPDT0601.Columns(9).Value)
 lArrDet(9) = CE(tdbgPDT0601.Columns(4).Value)
 lArrDet(10) = CE(tdbgPDT0601.Columns(5).Value)
        
 clsMante.InicializaClase
 clsMante.BeginTrans
        
 If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPDT0601", lArrDet(), False) = False Then
  Mensajes "El proceso no ha concluido....", vbInformation
  Screen.MousePointer = vbDefault
            
  clsMante.CancelTrans
  clsMante.FinalizaClase
            
  cmdEliminarTodo.Enabled = True
            
  Exit Sub
 End If
        
 clsMante.CommitTrans
 clsMante.FinalizaClase
 Set clsMante = Nothing
 
Exit Sub

Control:
 MsgBox Err.Description
 Screen.MousePointer = vbDefault
End Sub

Private Sub tdbgPDT0601_ColEdit(ByVal ColIndex As Integer)
If tdbgPDT0601.Col = 12 Then
    lArrDatos(tdbgPDT0601.Row, 12) = tdbgPDT0601.Columns(12).Value
End If

End Sub

Private Sub tdbgPDT0601_DblClick()
'PGBV 28102013
    Dim i As Integer
    Dim Contador As Integer
    Contador = lArrDatos.Count(1) - 1
    If tdbgPDT0601.DestinationCol = 12 Then
        For i = 0 To Contador
            If CE(lArrDatos(i, 1)) <> "" Then
                lArrDatos(i, 12) = "x"
            End If
        Next i
        tdbgPDT0601.Refresh
    End If
End Sub

Private Sub tdbgPDT0601_KeyPress(KeyAscii As Integer)
'PGBV 28102013
If KeyAscii = 13 And tdbgPDT0601.Col = 12 And tdbgPDT0601.Columns(0).Value <> "" Then
    tdbgPDT0601.Row = tdbgPDT0601.Row + 1
End If
End Sub
