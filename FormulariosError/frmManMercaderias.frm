VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManMercaderias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Mercaderias, Existencias y Suministros"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   Icon            =   "frmManMercaderias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   12570
   Begin TrueOleDBGrid70.TDBGrid grdCapital 
      Height          =   5235
      Left            =   90
      TabIndex        =   9
      Top             =   1260
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   9234
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Existencias"
      Columns(0).DataField=   ""
      Columns(0).DataWidth=   12
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Item"
      Columns(1).DataField=   "3"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Codigo"
      Columns(2).DataField=   ""
      Columns(2).DataWidth=   2
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo Existencia"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tdbdTipoExist"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Descripción"
      Columns(4).DataField=   ""
      Columns(4).DataWidth=   250
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "codigo"
      Columns(5).DataField=   ""
      Columns(5).DataWidth=   3
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Unidad de Medida"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6).DropDown=   "tdbdUMedida"
      Columns(6).DropDown.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Cantidad"
      Columns(7).DataField=   ""
      Columns(7).DefaultValue=   "0.00"
      Columns(7).DefaultValue.vt=   8
      Columns(7).NumberFormat=   "External Editor"
      Columns(7).ExternalEditor=   "TDBNumberNeg"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Costo Unitario"
      Columns(8).DataField=   ""
      Columns(8).DefaultValue=   "0"
      Columns(8).DefaultValue.vt=   8
      Columns(8).NumberFormat=   "External Editor"
      Columns(8).ExternalEditor=   "TDBNumberNegCosto"
      Columns(8).ExternalEditor.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Costo Total"
      Columns(9).DataField=   ""
      Columns(9).DefaultValue=   "0"
      Columns(9).DefaultValue.vt=   8
      Columns(9).NumberFormat=   "External Editor"
      Columns(9).ExternalEditor=   "TDBNumberNeg"
      Columns(9).ExternalEditor.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "FLAG"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Clase"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Cuenta"
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "FormatText Event"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Denominación"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Saldo Cont."
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "External Editor"
      Columns(14).ExternalEditor=   "TDBNumberNeg"
      Columns(14).ExternalEditor.vt=   8
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   4
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2434"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2355"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=197124"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=197124"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=4498"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=4419"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=197124"
      Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=2752"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=2672"
      Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=197120"
      Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(33)=   "Column(3).DropDownList=1"
      Splits(0)._ColumnProps(34)=   "Column(4).Width=5159"
      Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=5080"
      Splits(0)._ColumnProps(37)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=197124"
      Splits(0)._ColumnProps(39)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(41)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(42)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=197120"
      Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(50)=   "Column(6).Width=2619"
      Splits(0)._ColumnProps(51)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(6)._WidthInPix=2540"
      Splits(0)._ColumnProps(53)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(54)=   "Column(6)._ColStyle=197120"
      Splits(0)._ColumnProps(55)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(56)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(57)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(58)=   "Column(6).AutoDropDown=1"
      Splits(0)._ColumnProps(59)=   "Column(6).DropDownList=1"
      Splits(0)._ColumnProps(60)=   "Column(7).Width=1429"
      Splits(0)._ColumnProps(61)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(7)._WidthInPix=1349"
      Splits(0)._ColumnProps(63)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(64)=   "Column(7)._ColStyle=197122"
      Splits(0)._ColumnProps(65)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(67)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(68)=   "Column(8).Width=2143"
      Splits(0)._ColumnProps(69)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(8)._WidthInPix=2064"
      Splits(0)._ColumnProps(71)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(72)=   "Column(8)._ColStyle=197122"
      Splits(0)._ColumnProps(73)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(74)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(75)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(76)=   "Column(9).Width=767"
      Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=688"
      Splits(0)._ColumnProps(79)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(80)=   "Column(9)._ColStyle=205314"
      Splits(0)._ColumnProps(81)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(83)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(84)=   "Column(10).Width=1667"
      Splits(0)._ColumnProps(85)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(10)._WidthInPix=1588"
      Splits(0)._ColumnProps(87)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(88)=   "Column(10)._ColStyle=197124"
      Splits(0)._ColumnProps(89)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(90)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(91)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(92)=   "Column(11).Width=873"
      Splits(0)._ColumnProps(93)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(11)._WidthInPix=794"
      Splits(0)._ColumnProps(95)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(96)=   "Column(11)._ColStyle=197121"
      Splits(0)._ColumnProps(97)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(98)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(99)=   "Column(11).Merge=1"
      Splits(0)._ColumnProps(100)=   "Column(12).Width=1508"
      Splits(0)._ColumnProps(101)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(12)._WidthInPix=1429"
      Splits(0)._ColumnProps(103)=   "Column(12)._ColStyle=205316"
      Splits(0)._ColumnProps(104)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(105)=   "Column(13).Width=2461"
      Splits(0)._ColumnProps(106)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(13)._WidthInPix=2381"
      Splits(0)._ColumnProps(108)=   "Column(13)._ColStyle=197124"
      Splits(0)._ColumnProps(109)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(110)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(111)=   "Column(14).Width=1799"
      Splits(0)._ColumnProps(112)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(113)=   "Column(14)._WidthInPix=1720"
      Splits(0)._ColumnProps(114)=   "Column(14)._ColStyle=197122"
      Splits(0)._ColumnProps(115)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(116)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(117)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(121)=   "Column(15)._ColStyle=205316"
      Splits(0)._ColumnProps(122)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(123)=   "Column(15).Order=16"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).Size  =   8
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=16"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
      Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(9)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(10)=   "Column(1)._ColStyle=197124"
      Splits(1)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(14)=   "Column(2).Width=4498"
      Splits(1)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(16)=   "Column(2)._WidthInPix=4419"
      Splits(1)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(18)=   "Column(2)._ColStyle=197124"
      Splits(1)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(22)=   "Column(3).Width=2593"
      Splits(1)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(24)=   "Column(3)._WidthInPix=2514"
      Splits(1)._ColumnProps(25)=   "Column(3)._ColStyle=197120"
      Splits(1)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(27)=   "Column(3).DropDownList=1"
      Splits(1)._ColumnProps(28)=   "Column(4).Width=3784"
      Splits(1)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(30)=   "Column(4)._WidthInPix=3704"
      Splits(1)._ColumnProps(31)=   "Column(4)._ColStyle=197124"
      Splits(1)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(33)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(36)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(5)._ColStyle=197120"
      Splits(1)._ColumnProps(38)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(41)=   "Column(6).Width=2699"
      Splits(1)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(6)._WidthInPix=2619"
      Splits(1)._ColumnProps(44)=   "Column(6)._ColStyle=197120"
      Splits(1)._ColumnProps(45)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(46)=   "Column(6).AutoDropDown=1"
      Splits(1)._ColumnProps(47)=   "Column(6).DropDownList=1"
      Splits(1)._ColumnProps(48)=   "Column(7).Width=1905"
      Splits(1)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(50)=   "Column(7)._WidthInPix=1826"
      Splits(1)._ColumnProps(51)=   "Column(7)._ColStyle=197122"
      Splits(1)._ColumnProps(52)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(53)=   "Column(8).Width=1720"
      Splits(1)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(55)=   "Column(8)._WidthInPix=1640"
      Splits(1)._ColumnProps(56)=   "Column(8)._ColStyle=197122"
      Splits(1)._ColumnProps(57)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(58)=   "Column(9).Width=3704"
      Splits(1)._ColumnProps(59)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(60)=   "Column(9)._WidthInPix=3625"
      Splits(1)._ColumnProps(61)=   "Column(9)._ColStyle=197122"
      Splits(1)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(63)=   "Column(10).Width=2672"
      Splits(1)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(10)._WidthInPix=2593"
      Splits(1)._ColumnProps(66)=   "Column(10).AllowSizing=0"
      Splits(1)._ColumnProps(67)=   "Column(10)._ColStyle=197124"
      Splits(1)._ColumnProps(68)=   "Column(10).Visible=0"
      Splits(1)._ColumnProps(69)=   "Column(10).AllowFocus=0"
      Splits(1)._ColumnProps(70)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(71)=   "Column(11).Width=2725"
      Splits(1)._ColumnProps(72)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(11)._WidthInPix=2646"
      Splits(1)._ColumnProps(74)=   "Column(11).AllowSizing=0"
      Splits(1)._ColumnProps(75)=   "Column(11)._ColStyle=197124"
      Splits(1)._ColumnProps(76)=   "Column(11).Visible=0"
      Splits(1)._ColumnProps(77)=   "Column(11).AllowFocus=0"
      Splits(1)._ColumnProps(78)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(79)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(80)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(81)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(82)=   "Column(12).AllowSizing=0"
      Splits(1)._ColumnProps(83)=   "Column(12)._ColStyle=197124"
      Splits(1)._ColumnProps(84)=   "Column(12).Visible=0"
      Splits(1)._ColumnProps(85)=   "Column(12).AllowFocus=0"
      Splits(1)._ColumnProps(86)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(87)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(88)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(90)=   "Column(13).AllowSizing=0"
      Splits(1)._ColumnProps(91)=   "Column(13)._ColStyle=197124"
      Splits(1)._ColumnProps(92)=   "Column(13).Visible=0"
      Splits(1)._ColumnProps(93)=   "Column(13).AllowFocus=0"
      Splits(1)._ColumnProps(94)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(95)=   "Column(14).Width=1482"
      Splits(1)._ColumnProps(96)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(14)._WidthInPix=1402"
      Splits(1)._ColumnProps(98)=   "Column(14).AllowSizing=0"
      Splits(1)._ColumnProps(99)=   "Column(14)._ColStyle=197122"
      Splits(1)._ColumnProps(100)=   "Column(14).Visible=0"
      Splits(1)._ColumnProps(101)=   "Column(14).AllowFocus=0"
      Splits(1)._ColumnProps(102)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(103)=   "Column(15).Width=238"
      Splits(1)._ColumnProps(104)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(15)._WidthInPix=159"
      Splits(1)._ColumnProps(106)=   "Column(15).AllowSizing=0"
      Splits(1)._ColumnProps(107)=   "Column(15)._ColStyle=205316"
      Splits(1)._ColumnProps(108)=   "Column(15).Button=1"
      Splits(1)._ColumnProps(109)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(110)=   "Column(15).ButtonText=1"
      Splits(1)._ColumnProps(111)=   "Column(15).ButtonAlways=1"
      Splits(1)._ColumnProps(112)=   "Column(15)._MinWidth=3"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   3
      FootLines       =   1
      MultipleLines   =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=37,.bgcolor=&HFFDBBB&,.bold=0"
      _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=25,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=27,.parent=3,.alignment=1"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=43,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=94,.parent=25,.bgcolor=&HFFFFFF&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=26"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=27,.bgcolor=&HC0C0C0&"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=43"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=142,.parent=25"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=139,.parent=26"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=140,.parent=27"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=141,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=126,.parent=25,.bgcolor=&HFFFFFF&"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=123,.parent=26"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=124,.parent=27,.bgcolor=&HC0C0C0&"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=125,.parent=43"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=82,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(50)  =   ":id=82,.locked=0"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=26"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=27,.bgcolor=&HC0C0C0&"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=43"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=86,.parent=25,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=26"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=27,.bgcolor=&HC0C0C0&"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=43"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=150,.parent=25,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=147,.parent=26"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=148,.parent=27"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=149,.parent=43"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=90,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=26"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=27,.bgcolor=&HC0C0C0&"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=43"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=102,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=26"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=27,.bgcolor=&HF8ECC9&"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=43,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=106,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(71)  =   ":id=106,.locked=0"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=103,.parent=26"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=104,.parent=27,.alignment=1"
      _StyleDefs(74)  =   ":id=104,.bgcolor=&HF8ECC9&"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=105,.parent=43"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=114,.parent=25,.alignment=1,.bgcolor=&HF8ECC9&"
      _StyleDefs(77)  =   ":id=114,.locked=-1"
      _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=111,.parent=26"
      _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=112,.parent=27,.bgcolor=&HF8ECC9&"
      _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=113,.parent=43"
      _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=158,.parent=25"
      _StyleDefs(82)  =   "Splits(0).Columns(10).HeadingStyle:id=155,.parent=26"
      _StyleDefs(83)  =   "Splits(0).Columns(10).FooterStyle:id=156,.parent=27"
      _StyleDefs(84)  =   "Splits(0).Columns(10).EditorStyle:id=157,.parent=43"
      _StyleDefs(85)  =   "Splits(0).Columns(11).Style:id=162,.parent=25,.alignment=2,.bgcolor=&HC08000&"
      _StyleDefs(86)  =   ":id=162,.fgcolor=&HFFFFFF&"
      _StyleDefs(87)  =   "Splits(0).Columns(11).HeadingStyle:id=159,.parent=26"
      _StyleDefs(88)  =   "Splits(0).Columns(11).FooterStyle:id=160,.parent=27,.fgcolor=&HFFFFFF&"
      _StyleDefs(89)  =   "Splits(0).Columns(11).EditorStyle:id=161,.parent=43"
      _StyleDefs(90)  =   "Splits(0).Columns(12).Style:id=16,.parent=25,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(12).HeadingStyle:id=13,.parent=26"
      _StyleDefs(92)  =   "Splits(0).Columns(12).FooterStyle:id=14,.parent=27,.bgcolor=&HFFDBBB&"
      _StyleDefs(93)  =   "Splits(0).Columns(12).EditorStyle:id=15,.parent=43"
      _StyleDefs(94)  =   "Splits(0).Columns(13).Style:id=146,.parent=25,.locked=0"
      _StyleDefs(95)  =   "Splits(0).Columns(13).HeadingStyle:id=143,.parent=26"
      _StyleDefs(96)  =   "Splits(0).Columns(13).FooterStyle:id=144,.parent=27"
      _StyleDefs(97)  =   "Splits(0).Columns(13).EditorStyle:id=145,.parent=43"
      _StyleDefs(98)  =   "Splits(0).Columns(14).Style:id=20,.parent=25,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(14).HeadingStyle:id=17,.parent=26"
      _StyleDefs(100) =   "Splits(0).Columns(14).FooterStyle:id=18,.parent=27"
      _StyleDefs(101) =   "Splits(0).Columns(14).EditorStyle:id=19,.parent=43"
      _StyleDefs(102) =   "Splits(0).Columns(15).Style:id=170,.parent=25,.locked=-1"
      _StyleDefs(103) =   "Splits(0).Columns(15).HeadingStyle:id=167,.parent=26"
      _StyleDefs(104) =   "Splits(0).Columns(15).FooterStyle:id=168,.parent=27"
      _StyleDefs(105) =   "Splits(0).Columns(15).EditorStyle:id=169,.parent=43"
      _StyleDefs(106) =   "Splits(1).Style:id=21,.parent=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(107) =   "Splits(1).CaptionStyle:id=48,.parent=4"
      _StyleDefs(108) =   "Splits(1).HeadingStyle:id=22,.parent=2"
      _StyleDefs(109) =   "Splits(1).FooterStyle:id=23,.parent=3,.alignment=1"
      _StyleDefs(110) =   "Splits(1).InactiveStyle:id=24,.parent=5"
      _StyleDefs(111) =   "Splits(1).SelectedStyle:id=30,.parent=6"
      _StyleDefs(112) =   "Splits(1).EditorStyle:id=29,.parent=7"
      _StyleDefs(113) =   "Splits(1).HighlightRowStyle:id=31,.parent=8"
      _StyleDefs(114) =   "Splits(1).EvenRowStyle:id=32,.parent=9"
      _StyleDefs(115) =   "Splits(1).OddRowStyle:id=47,.parent=10"
      _StyleDefs(116) =   "Splits(1).RecordSelectorStyle:id=49,.parent=11"
      _StyleDefs(117) =   "Splits(1).FilterBarStyle:id=50,.parent=12"
      _StyleDefs(118) =   "Splits(1).Columns(0).Style:id=54,.parent=21,.bgcolor=&HFFFFFF&"
      _StyleDefs(119) =   "Splits(1).Columns(0).HeadingStyle:id=51,.parent=22"
      _StyleDefs(120) =   "Splits(1).Columns(0).FooterStyle:id=52,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(121) =   "Splits(1).Columns(0).EditorStyle:id=53,.parent=29"
      _StyleDefs(122) =   "Splits(1).Columns(1).Style:id=58,.parent=21"
      _StyleDefs(123) =   "Splits(1).Columns(1).HeadingStyle:id=55,.parent=22"
      _StyleDefs(124) =   "Splits(1).Columns(1).FooterStyle:id=56,.parent=23"
      _StyleDefs(125) =   "Splits(1).Columns(1).EditorStyle:id=57,.parent=29"
      _StyleDefs(126) =   "Splits(1).Columns(2).Style:id=62,.parent=21,.bgcolor=&HFFFFFF&"
      _StyleDefs(127) =   "Splits(1).Columns(2).HeadingStyle:id=59,.parent=22"
      _StyleDefs(128) =   "Splits(1).Columns(2).FooterStyle:id=60,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(129) =   "Splits(1).Columns(2).EditorStyle:id=61,.parent=29"
      _StyleDefs(130) =   "Splits(1).Columns(3).Style:id=66,.parent=21,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(131) =   ":id=66,.locked=0"
      _StyleDefs(132) =   "Splits(1).Columns(3).HeadingStyle:id=63,.parent=22"
      _StyleDefs(133) =   "Splits(1).Columns(3).FooterStyle:id=64,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(134) =   "Splits(1).Columns(3).EditorStyle:id=65,.parent=29"
      _StyleDefs(135) =   "Splits(1).Columns(4).Style:id=70,.parent=21,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(136) =   "Splits(1).Columns(4).HeadingStyle:id=67,.parent=22"
      _StyleDefs(137) =   "Splits(1).Columns(4).FooterStyle:id=68,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(138) =   "Splits(1).Columns(4).EditorStyle:id=69,.parent=29"
      _StyleDefs(139) =   "Splits(1).Columns(5).Style:id=74,.parent=21,.alignment=0"
      _StyleDefs(140) =   "Splits(1).Columns(5).HeadingStyle:id=71,.parent=22"
      _StyleDefs(141) =   "Splits(1).Columns(5).FooterStyle:id=72,.parent=23"
      _StyleDefs(142) =   "Splits(1).Columns(5).EditorStyle:id=73,.parent=29"
      _StyleDefs(143) =   "Splits(1).Columns(6).Style:id=98,.parent=21,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(144) =   "Splits(1).Columns(6).HeadingStyle:id=95,.parent=22"
      _StyleDefs(145) =   "Splits(1).Columns(6).FooterStyle:id=96,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(146) =   "Splits(1).Columns(6).EditorStyle:id=97,.parent=29"
      _StyleDefs(147) =   "Splits(1).Columns(7).Style:id=110,.parent=21,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(148) =   "Splits(1).Columns(7).HeadingStyle:id=107,.parent=22"
      _StyleDefs(149) =   "Splits(1).Columns(7).FooterStyle:id=108,.parent=23,.bgcolor=&HF8ECC9&"
      _StyleDefs(150) =   "Splits(1).Columns(7).EditorStyle:id=109,.parent=29,.alignment=1"
      _StyleDefs(151) =   "Splits(1).Columns(8).Style:id=118,.parent=21,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(152) =   ":id=118,.locked=0"
      _StyleDefs(153) =   "Splits(1).Columns(8).HeadingStyle:id=115,.parent=22"
      _StyleDefs(154) =   "Splits(1).Columns(8).FooterStyle:id=116,.parent=23,.alignment=1"
      _StyleDefs(155) =   ":id=116,.bgcolor=&HF8ECC9&"
      _StyleDefs(156) =   "Splits(1).Columns(8).EditorStyle:id=117,.parent=29"
      _StyleDefs(157) =   "Splits(1).Columns(9).Style:id=122,.parent=21,.alignment=1,.bgcolor=&HF8ECC9&"
      _StyleDefs(158) =   ":id=122,.locked=0"
      _StyleDefs(159) =   "Splits(1).Columns(9).HeadingStyle:id=119,.parent=22"
      _StyleDefs(160) =   "Splits(1).Columns(9).FooterStyle:id=120,.parent=23,.bgcolor=&HF8ECC9&"
      _StyleDefs(161) =   "Splits(1).Columns(9).EditorStyle:id=121,.parent=29"
      _StyleDefs(162) =   "Splits(1).Columns(10).Style:id=130,.parent=21"
      _StyleDefs(163) =   "Splits(1).Columns(10).HeadingStyle:id=127,.parent=22"
      _StyleDefs(164) =   "Splits(1).Columns(10).FooterStyle:id=128,.parent=23"
      _StyleDefs(165) =   "Splits(1).Columns(10).EditorStyle:id=129,.parent=29"
      _StyleDefs(166) =   "Splits(1).Columns(11).Style:id=166,.parent=21"
      _StyleDefs(167) =   "Splits(1).Columns(11).HeadingStyle:id=163,.parent=22"
      _StyleDefs(168) =   "Splits(1).Columns(11).FooterStyle:id=164,.parent=23"
      _StyleDefs(169) =   "Splits(1).Columns(11).EditorStyle:id=165,.parent=29"
      _StyleDefs(170) =   "Splits(1).Columns(12).Style:id=134,.parent=21"
      _StyleDefs(171) =   "Splits(1).Columns(12).HeadingStyle:id=131,.parent=22"
      _StyleDefs(172) =   "Splits(1).Columns(12).FooterStyle:id=132,.parent=23"
      _StyleDefs(173) =   "Splits(1).Columns(12).EditorStyle:id=133,.parent=29"
      _StyleDefs(174) =   "Splits(1).Columns(13).Style:id=154,.parent=21"
      _StyleDefs(175) =   "Splits(1).Columns(13).HeadingStyle:id=151,.parent=22"
      _StyleDefs(176) =   "Splits(1).Columns(13).FooterStyle:id=152,.parent=23"
      _StyleDefs(177) =   "Splits(1).Columns(13).EditorStyle:id=153,.parent=29"
      _StyleDefs(178) =   "Splits(1).Columns(14).Style:id=138,.parent=21,.alignment=1"
      _StyleDefs(179) =   "Splits(1).Columns(14).HeadingStyle:id=135,.parent=22"
      _StyleDefs(180) =   "Splits(1).Columns(14).FooterStyle:id=136,.parent=23,.bgcolor=&HC0C0C0&"
      _StyleDefs(181) =   "Splits(1).Columns(14).EditorStyle:id=137,.parent=29"
      _StyleDefs(182) =   "Splits(1).Columns(15).Style:id=174,.parent=21,.locked=-1"
      _StyleDefs(183) =   "Splits(1).Columns(15).HeadingStyle:id=171,.parent=22"
      _StyleDefs(184) =   "Splits(1).Columns(15).FooterStyle:id=172,.parent=23"
      _StyleDefs(185) =   "Splits(1).Columns(15).EditorStyle:id=173,.parent=29"
      _StyleDefs(186) =   "Named:id=33:Normal"
      _StyleDefs(187) =   ":id=33,.parent=0"
      _StyleDefs(188) =   "Named:id=34:Heading"
      _StyleDefs(189) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(190) =   ":id=34,.wraptext=-1"
      _StyleDefs(191) =   "Named:id=35:Footing"
      _StyleDefs(192) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(193) =   "Named:id=36:Selected"
      _StyleDefs(194) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(195) =   "Named:id=37:Caption"
      _StyleDefs(196) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(197) =   "Named:id=38:HighlightRow"
      _StyleDefs(198) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(199) =   "Named:id=39:EvenRow"
      _StyleDefs(200) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(201) =   "Named:id=40:OddRow"
      _StyleDefs(202) =   ":id=40,.parent=33"
      _StyleDefs(203) =   "Named:id=41:RecordSelector"
      _StyleDefs(204) =   ":id=41,.parent=34"
      _StyleDefs(205) =   "Named:id=42:FilterBar"
      _StyleDefs(206) =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText tdbtCuenta 
      Height          =   330
      Left            =   975
      TabIndex        =   19
      Top             =   4035
      Visible         =   0   'False
      Width           =   1800
      _Version        =   65536
      _ExtentX        =   3175
      _ExtentY        =   582
      Caption         =   "frmManMercaderias.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManMercaderias.frx":0F36
      Key             =   "frmManMercaderias.frx":0F54
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   0
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
      MaxLength       =   15
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
   Begin VB.Frame fraBotones 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      TabIndex        =   15
      Top             =   750
      Width           =   10950
      Begin MSForms.CommandButton cmdImportarTodos 
         Height          =   375
         Left            =   8745
         TabIndex        =   8
         ToolTipText     =   " Importa todos los movimientos "
         Top             =   45
         Visible         =   0   'False
         Width           =   2205
         Caption         =   "Importar del mes anterior"
         PicturePosition =   327683
         Size            =   "3889;661"
         Picture         =   "frmManMercaderias.frx":0F98
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminarTodo 
         Height          =   375
         Left            =   5805
         TabIndex        =   6
         ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
         Top             =   45
         Width           =   1380
         Caption         =   " Eliminar Todo"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":1532
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminaItem 
         Height          =   375
         Left            =   4365
         TabIndex        =   5
         ToolTipText     =   "Eliminar el movimientos seleccionado"
         Top             =   45
         Width           =   1380
         Caption         =   " Eliminar Item"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":1ACC
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsertarItem 
         Height          =   375
         Left            =   2925
         TabIndex        =   4
         ToolTipText     =   "Insertar el movimientos seleccionado"
         Top             =   45
         Width           =   1380
         Caption         =   " Insertar Mov."
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":2066
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGrabar 
         Height          =   375
         Left            =   1485
         TabIndex        =   3
         ToolTipText     =   "Grabar modificaciones"
         Top             =   45
         Width           =   1380
         Caption         =   " Grabar"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":2600
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   375
         Left            =   45
         TabIndex        =   2
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   45
         Width           =   1380
         Caption         =   " Listar"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":2B9A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   375
         Left            =   7245
         TabIndex        =   7
         Top             =   45
         Width           =   1380
         Caption         =   " Salir"
         PicturePosition =   131072
         Size            =   "2434;661"
         Picture         =   "frmManMercaderias.frx":3134
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin TrueOleDBGrid70.TDBDropDown tdbdUMedida 
      Height          =   1830
      Left            =   3840
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3228
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descripción"
      Columns(0).DataField=   "TAB_CDESCRIPCAMPO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "TAB_CCODIGO"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=6429"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6350"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=688"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=609"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   2
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   0
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
      DeadAreaBackColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
   Begin TrueOleDBGrid70.TDBDropDown tdbdTipoExist 
      Height          =   1830
      Left            =   3840
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3228
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descripción"
      Columns(0).DataField=   "TAB_CDESCRIPCAMPO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "TAB_CCODIGO"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=6668"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1270"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1191"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   2
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   0
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
      DeadAreaBackColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   5760
      TabIndex        =   1
      Top             =   90
      Width           =   2925
      _ExtentX        =   5159
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
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).DividerStyle=   2
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
      Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=688"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=609"
      Splits(0)._ColumnProps(10)=   "Column(1)._VertColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2117"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
      _PropDict       =   $"frmManMercaderias.frx":36CE
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
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
   Begin TDBNumber6Ctl.TDBNumber tdbNumber 
      Height          =   285
      Left            =   7560
      TabIndex        =   13
      Top             =   3645
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManMercaderias.frx":3755
      Caption         =   "frmManMercaderias.frx":3775
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManMercaderias.frx":37D9
      Keys            =   "frmManMercaderias.frx":37F7
      Spin            =   "frmManMercaderias.frx":3831
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
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
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumberNeg 
      Height          =   285
      Left            =   4470
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManMercaderias.frx":3859
      Caption         =   "frmManMercaderias.frx":3879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManMercaderias.frx":38DD
      Keys            =   "frmManMercaderias.frx":38FB
      Spin            =   "frmManMercaderias.frx":3935
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;-###,###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999
      MinValue        =   -999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TrueOleDBList70.TDBCombo tdbcMetodo 
      Height          =   300
      Left            =   1125
      TabIndex        =   0
      Top             =   90
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
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
      _PropDict       =   $"frmManMercaderias.frx":395D
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
   Begin TDBNumber6Ctl.TDBNumber TDBNumberNegCosto 
      Height          =   285
      Left            =   3255
      TabIndex        =   20
      Top             =   975
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManMercaderias.frx":39E4
      Caption         =   "frmManMercaderias.frx":3A04
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManMercaderias.frx":3A68
      Keys            =   "frmManMercaderias.frx":3A86
      Spin            =   "frmManMercaderias.frx":3AC0
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.00000;-#,###,###,##0.00000"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00000;-#,###,###,##0.00000"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999
      MinValue        =   -999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label2 
      Caption         =   "F12 : Copia Saldo Cont. a Costo Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   2
      Left            =   8835
      TabIndex        =   21
      Top             =   450
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "F11 : Calcula costo total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   1
      Left            =   8835
      TabIndex        =   18
      Top             =   255
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "F10 : Repetir denominación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   8835
      TabIndex        =   17
      Top             =   75
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "METODO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   135
      TabIndex        =   16
      Top             =   135
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4770
      TabIndex        =   12
      Top             =   135
      Width           =   780
   End
End
Attribute VB_Name = "frmManMercaderias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrCapital As New XArrayDB
Dim lArrCapitalAnt As New XArrayDB
Dim gsGrupo As String
Dim lArrDetalle(14) As Variant
Dim rsTipoExistencias As ADODB.Recordset
Dim rsUMedida As ADODB.Recordset
Dim gsSalirControl As Boolean 'PARA EL CONTROL TDBNUMBER QUE ESTA ASOCIADA A LA GRILLA DEL DETALLE
Dim gsColumna As Integer
Const gsColorBloqueado = &HFFDBBB
Const NUM_COL = 14

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


'Private Sub cmdBuscarMov_Click()
'    If tdbcMes.BoundText = "" Then
'        Mensajes "Seleccione el periodo"
'        pSetFocus tdbcMes
'        Exit Sub
'    End If
'
'    Dim sql As String
'    Dim rsVouchers As ADODB.Recordset
'
'    Call CerrarRecordSet(rsVouchers)
'
'    sql = "spCn_GrabaMercaderias 'BUSCAR_VOUCHER', '" & gsEmpresa & "','" & gsAnio & "','" & tdbcMes.BoundText & "','" & tdbcMetodo.BoundText & "'"
'    Call LlenarRecordSet(sql, rsVouchers)
'
'    Dim Filas As Integer
'    Filas = CuentaFilas 'lArrCapital.Count(1)
'
'    On Error GoTo Siguiente
'    If Filas = 1 And CE(lArrCapital(0, 0)) = "" Then Filas = 0
'
'Siguiente:
'    If Filas < 0 Then Filas = 0
'    If Not rsVouchers Is Nothing Then
'       Do While Not rsVouchers.EOF
'          If rsVouchers.RecordCount <= 0 Then
'             Mensajes "No se encontraron vouchers con las cuenta 2 para este mes"
'             Exit Sub
'          End If
'
''          If BuscaEntidadMercaderia(CE(rsVouchers.Fields("ASE_NVOUCHER")), _
''                          NE(rsVouchers.Fields("ASD_NITEM"))) = False Then
'
'
'                lArrCapital.ReDim 0, Filas, 0, NUM_COL   ' filas
'
''                lArrCapital(Filas, 0) = CE(rsVouchers.Fields("ASE_NVOUCHER"))
''                lArrCapital(Filas, 1) = CE(rsVouchers.Fields("Ase_cNumMov"))
''                lArrCapital(Filas, 2) = CE(rsVouchers.Fields("Pla_cCuentacontable"))
''                lArrCapital(Filas, 3) = CE(rsVouchers.Fields("ASD_NITEM"))
''                lArrCapital(Filas, 4) = CE(rsVouchers.Fields("Asd_dFecDoc"))
''                lArrCapital(Filas, 5) = CE(rsVouchers.Fields("ASD_CGLOSA")) 'tipo de titulo
''                lArrCapital(Filas, 6) = "" 'codigo intang
'                lArrCapital(Filas, 7) = "0.00" 'cantidad
'                lArrCapital(Filas, 8) = "0.00" 'costo unitario
'                lArrCapital(Filas, 9) = "0.00" 'total
''                lArrCapital(Filas, 10) = NE(rsVouchers.Fields("IMPORTE"))   'val neto
'                lArrCapital(Filas, 11) = CE(rsVouchers.Fields("d2"))
'                lArrCapital(Filas, 12) = CE(rsVouchers.Fields("Pla_cCuentacontable"))
'                lArrCapital(Filas, 13) = CE(rsVouchers.Fields("pla_cnombrecuenta"))
'                lArrCapital(Filas, 14) = NE(rsVouchers.Fields("IMPORTE"))
'
'                Filas = Filas + 1
''          End If
'
'          rsVouchers.MoveNext
'       Loop
'       grdCapital.ReBind
'
'
'    Else
'       Mensajes "No se encontraron vouchers con la clase 2"
'    End If
'
'End Sub

Private Function BuscaEntidadMercaderia(Voucher As String, item As Integer) As Boolean
    BuscaEntidadMercaderia = False
    Dim i As Integer
    On Error GoTo serror
    If (lArrCapital.Count(1) = 1 Or lArrCapital.Count(2) = 1) And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
       BuscaEntidadMercaderia = False
       Exit Function
    End If
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 0)) = Voucher And _
           NE(lArrCapital(i, 3)) = item Then
           
           BuscaEntidadMercaderia = True
           Exit For
        End If
    Next i
    Exit Function
serror:
    BuscaEntidadMercaderia = False
End Function

Private Sub cmdEliminaItem_Click()
   If MsgBox("Deseas eliminar la fila seleccionada", vbYesNo + vbQuestion) = vbYes Then
      On Error Resume Next
      lArrCapital.DeleteRows grdCapital.Bookmark
      DoEvents
      grdCapital.Update
      DoEvents
      grdCapital.ReBind
   End If
End Sub

Private Function CuentaFilas() As Integer
    Dim i As Integer
    Dim Contador As Integer
    Contador = 0
    
    If lArrCapital.Count(1) = 1 And lArrCapital.Count(2) = 1 Then
        lArrCapital.ReDim 0, 0, 0, NUM_COL ' filas
    Else
        For i = 0 To lArrCapital.Count(1) - 1
            If CE(lArrCapital(i, 12)) <> "" Then
               Contador = Contador + 1
            End If
        Next i
    
    End If
    
    CuentaFilas = Contador
End Function

Private Sub cmdEliminarTodo_Click()
    If MsgBox("Deseas eliminar todas las filas de la lista", vbYesNo + vbQuestion) = vbYes Then
       lArrCapital.ReDim 0, 0, 0, NUM_COL ' filas
       lArrCapital.Clear
       DoEvents
       grdCapital.Update
       DoEvents
       grdCapital.ReBind
    End If
End Sub

Private Sub cmdGrabar_Click()
    Call Grabar
End Sub

Private Sub AgregaFila()
    Dim Filas As Integer
    Filas = CuentaFilas
    lArrCapital.ReDim 0, Filas, 0, NUM_COL   ' filas
    
    lArrCapital(Filas, 7) = "0.00"
    lArrCapital(Filas, 8) = "0.00"
    lArrCapital(Filas, 9) = "0.00"
    
    Call UpdateGrilla
    grdCapital.ReBind

End Sub

Private Sub cmdImportarTodos_Click()
    Call GeneraArregloAnt
End Sub

Private Sub cmdInsertarItem_Click()
    If tdbcMes.Text = "" Then
        'Mensajes "Seleccione el mes a ingresar"
        pSetFocus tdbcMes
        Exit Sub
    End If

    Call AgregaFila
    
End Sub

Private Function BuscaEntidad(Voucher As String, item As Integer, Tipo As String, Codigo As String) As Boolean
    BuscaEntidad = False
    Dim i As Integer
    On Error GoTo serror
    If (lArrCapital.Count(1) = 1 Or lArrCapital.Count(2) = 1) And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
       BuscaEntidad = False
       Exit Function
    End If
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 0)) = Voucher And _
            NE(lArrCapital(i, 1)) = item And _
            CE(lArrCapital(i, 2)) = Tipo And _
            CE(lArrCapital(i, 3)) = Codigo Then
                BuscaEntidad = True
                Exit For
        End If
    Next i
    
    Exit Function
serror:
    BuscaEntidad = False
End Function

Private Sub cmdRefresh_Click()
    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    
    GeneraArreglo
    DoEvents
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
    
End Sub

Private Function BuscaImporte(ByRef rs As ADODB.Recordset, Voucher As String, item As Integer, Tipo As String, Codigo As String) As Double
    BuscaImporte = -1
    
    If Not rs Is Nothing Then
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
               If rs.Fields("ASE_NVOUCHER") = Voucher And _
                  rs.Fields("ASD_NITEM") = item And _
                  rs.Fields("TEN_CTIPOENTIDAD") = Tipo And _
                  rs.Fields("ENT_CCODENTIDAD") = Codigo Then
                  
                  BuscaImporte = NE(rs.Fields("IMPORTE"))
                  Exit Do
               End If
               rs.MoveNext
            Loop
        End If
    End If
End Function

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdVerificar_Click()
'    If lArrCapital.Count(2) = 1 Then
'        Exit Sub
'    End If
'
'    Dim Sql As String
'    Dim rsVouchers As ADODB.Recordset
'
'    cmdVerificar.Enabled = False
'    Screen.MousePointer = vbHourglass
'    DoEvents
'
'    Call CerrarRecordSet(rsVouchers)
'
'    Sql = "spCn_RptFormato0308 'VERIFICAR_VOUCHER', '" & gsEmpresa & "','" & gsAnio & "','" & tdbcMes.BoundText & "'"
'    Call LlenarRecordSet(Sql, rsVouchers)
'
'    Dim i As Integer
'    Dim Importe As Double
'    For i = 0 To lArrCapital.Count(1) - 1
'        Importe = BuscaImporte(rsVouchers, _
'                               CE(lArrCapital(i, 0)), _
'                               NE(lArrCapital(i, 1)), _
'                               CE(lArrCapital(i, 2)), _
'                               CE(lArrCapital(i, 3)))
'
'        If Importe = NE(lArrCapital(i, 11)) Then lArrCapital(i, 12) = "0"
'        If Importe <> NE(lArrCapital(i, 11)) Then lArrCapital(i, 12) = "1"
'        If Importe = -1 Then lArrCapital(i, 12) = "2"
'        If CE(lArrCapital(i, 0)) = "" Then lArrCapital(i, 12) = "0"
'    Next i
'
'    grdCapital.Refresh
'
'    cmdVerificar.Enabled = True
'    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
                If MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar") = vbYes Then
                    Unload Me
                End If

        Case 115: If cmdGrabar.Enabled Then cmdGrabar_Click
        'Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        'Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select


End Sub

Private Sub CargarCombos()
    Dim sqlcombos  As String
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    DoEvents
    tdbcMes.BoundText = gsPeriodo
    tdbcMes.ReBind
    DoEvents
    '---------------------------------------------------------
    Call CerrarRecordSet(rsTipoExistencias)
    
    sqlcombos = "SELECT TAB_CDESCRIPCAMPO , TAB_CCODIGO " & _
                "FROM TABLA WHERE TAB_CTABLA='080' AND EMP_CCODIGO='" & gsEmpresa & "' " & _
                "ORDER BY TAB_CCODIGO "
    
    Call LlenarRecordSet(sqlcombos, rsTipoExistencias)
    Set tdbdTipoExist.DataSource = rsTipoExistencias
    '------------------------------------------------------------
    Call CerrarRecordSet(rsUMedida)
    
    sqlcombos = "SELECT TAB_CDESCRIPCAMPO , TAB_CCODIGO " & _
                "FROM TABLA WHERE TAB_CTABLA='053' AND EMP_CCODIGO='" & gsEmpresa & "' " & _
                "ORDER BY TAB_CCODIGO "
    
    Call LlenarRecordSet(sqlcombos, rsUMedida)
    Set tdbdUMedida.DataSource = rsUMedida
    '------------------------------------------------------------
    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA WITH(NOLOCK) " & _
                "WHERE Emp_cCodigo='" & gsEmpresa & "' AND Tab_cTabla = '085' " & _
                "ORDER BY Tab_cCodigo"
                
    LlenarComboAddItem tdbcMetodo, sqlcombos
    

End Sub

Private Sub SumarTotales()
    Dim i As Integer
    Dim iFila As Integer
    
    
    Dim s_Cantidad As Double, s_Unitario As Double
    Dim s_Total As Double
    
    On Error GoTo serror
    iFila = lArrCapital.Count(1)
    
    For i = 0 To iFila - 1
        s_Cantidad = s_Cantidad + NE(lArrCapital.Value(i, 7))
        s_Unitario = s_Unitario + NE(lArrCapital.Value(i, 8))
        s_Total = s_Total + NE(lArrCapital.Value(i, 9))
    Next i

    grdCapital.Columns(7).FooterText = Format(s_Cantidad, "###,###,##0.00")
    grdCapital.Columns(8).FooterText = Format(s_Unitario, "###,###,##0.00")
    grdCapital.Columns(9).FooterText = Format(s_Total, "###,###,##0.00")
    
    Exit Sub
    
serror:
    s_Cantidad = 0
    s_Unitario = 0
    s_Total = 0
    
    grdCapital.Columns(7).FooterText = Format(s_Cantidad, "###,###,##0.00")
    grdCapital.Columns(8).FooterText = Format(s_Unitario, "###,###,##0.00")
    grdCapital.Columns(9).FooterText = Format(s_Total, "###,###,##0.00")

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    Centrar_form Me
    
    Me.Height = 6990
    Me.Width = 11550
    '-----------------------------
    grdCapital.Splits(0).MarqueeStyle = dbgSolidCellBorder
    Call GeneraArreglo
    DoEvents
    Call AgregaFila
    Call SumarTotales
    Call CargarCombos
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabar.Enabled = False
        cmdInsertarItem.Enabled = False
        cmdEliminaItem.Enabled = False
        cmdEliminarTodo.Enabled = False
        grdCapital.Splits(0).Locked = True
    Else
        cmdGrabar.Enabled = True
        cmdInsertarItem.Enabled = True
        cmdEliminaItem.Enabled = True
        cmdEliminarTodo.Enabled = True
        grdCapital.Splits(0).Locked = False
    End If
    
    
End Sub

Private Sub GeneraArreglo()
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion
    
    sql = "spCn_GrabaMercaderias 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "', '" & tdbcMetodo.BoundText & "'"
    Call GridArreglo(lArrCapital, grdCapital, sql)
    
    grdCapital.Splits(1).ScrollBars = dbgBoth
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub

Private Function BuscaFila(scadena As String) As Boolean
    Dim i As Integer
    BuscaFila = False
    Dim sFila As String

    On Error GoTo ErrorEjecucion
    
    For i = 0 To lArrCapital.Count(1) - 1
        sFila = CE(lArrCapital(i, 0)) & _
                CE(lArrCapital(i, 2)) & _
                CE(lArrCapital(i, 4)) & _
                CE(lArrCapital(i, 5))
        
        If sFila = scadena Then
            BuscaFila = True
            Exit Function
        End If
    Next i
    
    
    Exit Function
ErrorEjecucion:
    'Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
    
End Function


Private Sub GeneraArregloAnt()

    Dim sql As String
    Dim MesAnt As String, scadena As String
    Dim i As Integer
    Dim k As Integer
    Dim nI As Integer
    Dim Fila As Integer
    Dim entro As Boolean
    On Local Error GoTo ErrorEjecucion
    entro = False
    nI = 0
    MesAnt = Right("00" & CE(Val(tdbcMes.BoundText) - 1), 2)
    
    Set lArrCapitalAnt = Nothing
    
    sql = "spCn_GrabaMercaderias 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & MesAnt & "', '" & tdbcMetodo.BoundText & "'"
    Call LlenarArreglo(lArrCapitalAnt, sql)
    
    If lArrCapitalAnt.Count(1) = 1 And lArrCapitalAnt.Count(2) = 1 Then
        Mensajes "No se encontraron datos en el mes anterior"
        Exit Sub
    End If
    
    Fila = CuentaFilas
   
    For i = 0 To lArrCapitalAnt.Count(1) - 1
        scadena = CE(lArrCapitalAnt(i, 0)) & _
                  CE(lArrCapitalAnt(i, 2)) & _
                  CE(lArrCapitalAnt(i, 4)) & _
                  CE(lArrCapitalAnt(i, 5))
    
        If BuscaFila(scadena) = False Then
            lArrCapital.ReDim 0, Fila + nI, 0, lArrCapitalAnt.Count(2) - 1
            
            For k = 0 To lArrCapital.Count(2) - 1
                lArrCapital(Fila + nI, k) = lArrCapitalAnt(i, k)
                entro = True
            Next k
            
            nI = nI + 1
        End If
        
    Next i
    
    If entro = False Then Mensajes "Datos del mes anterior ya fueron importados"
    
    Set grdCapital.Array = lArrCapital
    grdCapital.ReBind
    
    Exit Sub
ErrorEjecucion:
    'Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        'fraEstructura.Height = Me.Height - 1900 + 250
        'fraEstructura.Width = Me.Width - 150
        grdCapital.Height = Me.Height - 1700
        grdCapital.Width = Me.Width - 300
        grdCapital.Splits(1).ScrollBars = dbgNone
        grdCapital.Splits(1).ScrollBars = dbgAutomatic
   '     fraBotones.Top = grdCapital.Top + grdCapital.Height
'        fraLeyenda.Width = grdCapital.Width
    End If
    
    Exit Sub
    
serror:
    Mensajes Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub grdCapital_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 7 Then
       grdCapital.Columns(ColIndex) = NE(grdCapital.Columns(ColIndex).Value)
       
       grdCapital.Columns(9).Value = NE(grdCapital.Columns(7).Value) * NE(grdCapital.Columns(8).Value)
       pSetFocus grdCapital
    End If
    
    If ColIndex = 8 Then
       grdCapital.Columns(ColIndex) = Abs(NE(grdCapital.Columns(ColIndex).Value))
       
       grdCapital.Columns(9).Value = NE(grdCapital.Columns(7).Value) * NE(grdCapital.Columns(8).Value)
       pSetFocus grdCapital
    End If
    
End Sub

Private Sub grdCapital_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If tdbcMes.Text = "" Then
       Cancel = 1
    End If
    


'    If ColIndex = 7 Or ColIndex = 8 Then
'        If CE(lArrCapital(grdCapital.Bookmark, 13)) = "S" Then
'           Cancel = 1
'        End If
'    End If
End Sub

Private Sub UpdateGrilla()
    On Error Resume Next
    DoEvents
    grdCapital.Update
    DoEvents
End Sub


Private Sub grdCapital_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
    On Error GoTo serror
'    If Split = 1 Then
'        If Col = 7 Or Col = 8 Then
'            If lArrCapital.Count(2) > 1 Then
'                If CE(lArrCapital(Bookmark, 13)) = "S" Then
'                    CellStyle.BackColor = gsColorBloqueado
'                End If
'            End If
'        End If
'    End If
    Exit Sub
serror:
End Sub

Private Sub grdCapital_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
'    If lArrCapital Is Nothing Or IsNull(grdCapital.Bookmark) Then
'        Exit Sub
'    End If
'
'    On Error GoTo SERROR
'
'    If lArrCapital(Bookmark, 12) = "2" Then
'        RowStyle.BackColor = &HFF&
'        RowStyle.ForeColor = &HFFFF&
'    ElseIf lArrCapital(Bookmark, 12) = "1" Then
'        RowStyle.BackColor = gsColorDesactProv
'    ElseIf lArrCapital(Bookmark, 12) = "3" Then
'        RowStyle.BackColor = &HFFFFC0
'    End If
'
'    Exit Sub
'SERROR:
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
    Select Case lControl
           Case "CuentasFilt"
                grdCapital.Columns(grdCapital.Col).Value = param0
           
           Case "Cuentas"
                grdCapital.Columns(grdCapital.Col).Value = param0
                grdCapital.Columns(grdCapital.Col + 1).Value = param1
    End Select
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    
    If lArrCapital.Count(1) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
       ValidaCampos = True
       Exit Function
    End If
    
    Dim i As Integer
    
    For i = 0 To lArrCapital.Count(1) - 1
        If NE(lArrCapital(i, 8)) <> 0 Or CE(lArrCapital(i, 2)) <> "" Or CE(lArrCapital(i, 0)) <> "" Then
        
            If CE(lArrCapital(i, 0)) = "" Then
               Mensajes "Ingrese la descripcion de la EXISTENCIA "
               grdCapital.Bookmark = i
               grdCapital.Col = 0
               pSetFocus grdCapital
               Exit Function
            End If
        
        
            If CE(lArrCapital(i, 3)) = "" Then
               Mensajes "Ingrese el TIPO DE EXISTENCIA "
               grdCapital.Bookmark = i
               grdCapital.Col = 3
               pSetFocus grdCapital
               Exit Function
            End If
            
            If CE(lArrCapital(i, 6)) = "" Then
               Mensajes "Ingrese la UNIDAD DE MEDIDA"
               grdCapital.Bookmark = i
               grdCapital.Col = 6
               pSetFocus grdCapital
               Exit Function
            End If
        End If
        
    Next i
    ValidaCampos = True
End Function


Private Sub Grabar()
    On Error GoTo serror
    UpdateGrilla
    
    If lArrCapital.Count(2) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
        Exit Sub
    End If

    If ValidaCampos = False Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'UpdateGrilla

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas
    grdCapital.Bookmark = grdCapital.Bookmark

    Dim lArrDet(15) As Variant
    lArrDet(0) = "ELIMINAR" '    @Accion         varchar(20)= '',
    lArrDet(1) = gsEmpresa  '@Emp_cCodigo    char(3)='',
    lArrDet(2) = gsAnio  '@Pan_cAnio      char(4)='',
    lArrDet(3) = tdbcMes.BoundText   '@Per_cPeriodo   char(2)='',
    lArrDet(4) = tdbcMetodo.BoundText
    clsMante.InicializaClase
    clsMante.BeginTrans
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaMercaderias", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        
        clsMante.CancelTrans
        clsMante.FinalizaClase
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    lArrDet(0) = "INSERTAR" '    @Accion         varchar(20)= '',
    
    For i = 0 To lArrCapital.Count(1) - 1
    
        lArrDet(5) = Right("000" & CE(i + 1), 3) '@Mer_cItem      char(3)='',
        lArrDet(6) = CE(lArrCapital(i, 0)) '@Mer_cCodigo    varchar(12)='' ,
        lArrDet(7) = CE(lArrCapital(i, 2)) '@Mer_cTipo      char(2)='' ,
        lArrDet(8) = CE(lArrCapital(i, 4)) '@Mer_cDescrip   varchar(250)='' ,
        lArrDet(9) = CE(lArrCapital(i, 5)) '@Mer_cMedida    char(3)='' ,
        lArrDet(10) = NE(lArrCapital(i, 7)) '@Mer_nCantidad  numeric(14,2)=0 ,
        lArrDet(11) = NE(lArrCapital(i, 8)) '@Mer_nCosto     numeric(14,2)=0 ,
        lArrDet(12) = NE(lArrCapital(i, 9)) '@Mer_nTotal     numeric(14,2)=0 ,
        lArrDet(13) = CE(lArrCapital(i, 12))  '@Pla_cCuentaContable
        lArrDet(14) = gsUsuario  '@Mer_cUsuario
        lArrDet(15) = ""  '@Mer_cMoneda
        
        If CE(lArrCapital(i, 0)) <> "" Then
            
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaMercaderias", lArrDet(), False) = False Then
                Screen.MousePointer = vbNormal
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                
                clsMante.CancelTrans
                clsMante.FinalizaClase
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            
        End If
    Next

    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal

    Set clsMante = Nothing
    Call GeneraArreglo
    
    DoEvents
    
    cmdRefresh_Click
    
    DoEvents
    Mensajes "Se ha grabado con exito ...", vbInformation
    
    Exit Sub
serror:
    Mensajes Err.Description
End Sub


'Private Function CargaArregloDet(item As Integer) As Boolean
'    CargaArregloDet = True
'
'    lArrDetalle(0) = "INSERTAR"
'    lArrDetalle(1) = gsEmpresa
'    lArrDetalle(2) = gsAnio
'    lArrDetalle(3) = tdbcMes.BoundText
'    lArrDetalle(4) = CE(lArrCapital(item, 0)) 'voucher
'    lArrDetalle(5) = NE(lArrCapital(item, 1)) 'item
'    lArrDetalle(6) = CE(lArrCapital(item, 2)) 'tipo entidad
'    lArrDetalle(7) = CE(lArrCapital(item, 3)) 'cod entidad
'    lArrDetalle(8) = CE(lArrCapital(item, 5)) 'cod titulo
'    lArrDetalle(9) = CE(lArrCapital(item, 6)) 'desc titulo
'    lArrDetalle(10) = NE(lArrCapital(item, 7)) 'valor nom unit
'    lArrDetalle(11) = NE(lArrCapital(item, 8)) 'cantidad
'    lArrDetalle(12) = NE(lArrCapital(item, 9)) 'costo total
'    lArrDetalle(13) = NE(lArrCapital(item, 10)) 'prov total
'    lArrDetalle(14) = NE(lArrCapital(item, 11)) 'total neto
'
'End Function

Private Sub Imprimir()
    Screen.MousePointer = vbHourglass

    Dim matriz_fecha(4) As Variant
    Dim Tipo As String

    matriz_fecha(0) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(1) = "@RUC;" & gsRUC & ";True"
    matriz_fecha(2) = "@Tipo;BUSCARTODOS;True"
    matriz_fecha(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(4) = "@Pan_cAnio;" & gsAnio & ";True"

    Dim formulas(0) As Variant
    'AbreReporteParam gsDSN, Me, rutaReportes & "RptPatrimonioNeto.rpt", crptToWindow, "Reporte de Patrimonio Neto", "", matriz_fecha(), Formulas()
    Screen.MousePointer = vbDefault

End Sub

Private Sub grdCapital_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Then
        lArrCapital(grdCapital.Bookmark, 4) = lArrCapital(grdCapital.Bookmark, 13)
        KeyCode = 0
        grdCapital.Refresh
        grdCapital.SetFocus
    End If
    
    If KeyCode = vbKeyF11 Then
        UpdateGrilla
        If NE(lArrCapital(grdCapital.Bookmark, 7)) = 0 Then
            lArrCapital(grdCapital.Bookmark, 8) = 0
        Else
            
            lArrCapital(grdCapital.Bookmark, 8) = Round(NE(lArrCapital(grdCapital.Bookmark, 14)) / NE(lArrCapital(grdCapital.Bookmark, 7)), 5)
        End If
        
        grdCapital.Refresh
        
        Call CalculaCostoTotal
        KeyCode = 0
        grdCapital.Refresh
        grdCapital.SetFocus
    End If
    
    If KeyCode = vbKeyF12 Then
        Dim nCalculo As Double
        If NE(lArrCapital(grdCapital.Bookmark, 7)) = 0 Then
            nCalculo = 0
        Else
            
            nCalculo = Round(NE(lArrCapital(grdCapital.Bookmark, 14)) / NE(lArrCapital(grdCapital.Bookmark, 7)), 5)
        End If
        
        
        If nCalculo = 0 Then
            lArrCapital(grdCapital.Bookmark, 9) = 0
        Else
            lArrCapital(grdCapital.Bookmark, 9) = lArrCapital(grdCapital.Bookmark, 14)
        End If
        
        grdCapital.Refresh
        'UpdateGrilla
        

        KeyCode = 0
        grdCapital.Refresh
        grdCapital.SetFocus
    End If
    
End Sub

Private Sub grdCapital_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If lArrCapital.Count(2) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
        Exit Sub
    End If

    gsSalirControl = False
    'Call CalculaCostoTotal
    grdCapital.Update

    SumarTotales
    pSetFocus grdCapital
End Sub

Private Sub CalculaCostoTotal()
    UpdateGrilla
    grdCapital.Columns(9).Value = Redondear(NE(grdCapital.Columns(7).Value) * NE(grdCapital.Columns(8).Value), 2)
End Sub

Private Sub tdbcMes_ItemChange()
    cmdRefresh_Click
    cmdInsertarItem_Click
    DoEvents
    pSetFocus tdbcMes
End Sub

Private Sub tdbcMetodo_ItemChange()
    tdbcMes_ItemChange
    pSetFocus tdbcMetodo
End Sub

Private Sub tdbdTipoExist_DropDownClose()
    grdCapital.Columns(2).Value = tdbdTipoExist.Columns(1).Value
    DoEvents
    
    grdCapital.RefetchRow
    pSetFocus grdCapital
    DoEvents
End Sub

Private Sub tdbdUMedida_DropDownClose()
    grdCapital.Columns(5).Value = tdbdUMedida.Columns(1).Value
    DoEvents
    
    grdCapital.RefetchRow
    pSetFocus grdCapital
    DoEvents
End Sub

Private Sub tdbNumber_GotFocus()
    On Error Resume Next
    TDBNumber.Value = Abs(NE(lArrCapital(grdCapital.Bookmark, grdCapital.Col)))
End Sub

Private Sub tdbNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       ControlAbs TDBNumber
       pSetFocus grdCapital
    End If
End Sub

Private Sub tdbNumber_KeyPress(KeyAscii As Integer)
    If gsSalirControl = False Then
        gsSalirControl = True
        TDBNumber = "0.00"
    End If
End Sub

