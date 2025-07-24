VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManPatrimonioNeto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patrimonio Neto"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "frmPatrimonioNeto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10905
   Begin TrueOleDBGrid70.TDBGrid grdPatrimonio 
      Height          =   4485
      Left            =   90
      TabIndex        =   0
      Top             =   1110
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   7911
      _LayoutType     =   4
      _RowHeight      =   26
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cuentas Patrimoniales"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Capital"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "External Editor"
      Columns(2).ExternalEditor=   "TDBNumLite"
      Columns(2).ExternalEditor.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Capital Adicional"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "External Editor"
      Columns(3).ExternalEditor=   "TDBNumLite"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Acciones de Inversión"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "External Editor"
      Columns(4).ExternalEditor=   "TDBNumLite"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Excedente de Revaluación"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "External Editor"
      Columns(5).ExternalEditor=   "TDBNumLite"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Reserva Legal"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "External Editor"
      Columns(6).ExternalEditor=   "TDBNumLite"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Otras Reservas"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "External Editor"
      Columns(7).ExternalEditor=   "TDBNumLite"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Resultados Acumulados"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "External Editor"
      Columns(8).ExternalEditor=   "TDBNumLite"
      Columns(8).ExternalEditor.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=139777"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=9128"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=9049"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=139780"
      Splits(0)._ColumnProps(10)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2487"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2408"
      Splits(0)._ColumnProps(15)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=131586"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2778"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2699"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=131586"
      Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(26)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(28)=   "Column(4).Width=2699"
      Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=2619"
      Splits(0)._ColumnProps(31)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=131586"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=2778"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2699"
      Splits(0)._ColumnProps(39)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=131586"
      Splits(0)._ColumnProps(41)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(44)=   "Column(6).Width=2302"
      Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=2223"
      Splits(0)._ColumnProps(47)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=131586"
      Splits(0)._ColumnProps(49)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(51)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(52)=   "Column(7).Width=2514"
      Splits(0)._ColumnProps(53)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(7)._WidthInPix=2434"
      Splits(0)._ColumnProps(55)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._ColStyle=131586"
      Splits(0)._ColumnProps(57)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(59)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(60)=   "Column(8).Width=1429"
      Splits(0)._ColumnProps(61)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(8)._WidthInPix=1349"
      Splits(0)._ColumnProps(63)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(64)=   "Column(8)._ColStyle=131586"
      Splits(0)._ColumnProps(65)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(67)=   "Column(8).Order=9"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=9"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1085"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=139777"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=9128"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=9049"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=139780"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=2487"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=2408"
      Splits(1)._ColumnProps(20)=   "Column(2)._ColStyle=131584"
      Splits(1)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(22)=   "Column(3).Width=2778"
      Splits(1)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(24)=   "Column(3)._WidthInPix=2699"
      Splits(1)._ColumnProps(25)=   "Column(3)._ColStyle=131584"
      Splits(1)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(27)=   "Column(4).Width=2699"
      Splits(1)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(4)._WidthInPix=2619"
      Splits(1)._ColumnProps(30)=   "Column(4)._ColStyle=131584"
      Splits(1)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(32)=   "Column(5).Width=2778"
      Splits(1)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(34)=   "Column(5)._WidthInPix=2699"
      Splits(1)._ColumnProps(35)=   "Column(5)._ColStyle=131584"
      Splits(1)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(37)=   "Column(6).Width=2302"
      Splits(1)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(39)=   "Column(6)._WidthInPix=2223"
      Splits(1)._ColumnProps(40)=   "Column(6)._ColStyle=131584"
      Splits(1)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(42)=   "Column(7).Width=2514"
      Splits(1)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(44)=   "Column(7)._WidthInPix=2434"
      Splits(1)._ColumnProps(45)=   "Column(7)._ColStyle=131584"
      Splits(1)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(47)=   "Column(8).Width=1429"
      Splits(1)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(49)=   "Column(8)._WidthInPix=1349"
      Splits(1)._ColumnProps(50)=   "Column(8)._ColStyle=131584"
      Splits(1)._ColumnProps(51)=   "Column(8).Order=9"
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
      FootLines       =   0
      MultipleLines   =   0
      CellTips        =   2
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
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=37,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=98,.parent=13,.alignment=2,.bgcolor=&HF8ECC9&"
      _StyleDefs(38)  =   ":id=98,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HF8ECC9&,.wraptext=-1"
      _StyleDefs(43)  =   ":id=32,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=74,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(75)  =   "Splits(1).Style:id=25,.parent=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(76)  =   "Splits(1).CaptionStyle:id=76,.parent=4"
      _StyleDefs(77)  =   "Splits(1).HeadingStyle:id=26,.parent=2"
      _StyleDefs(78)  =   "Splits(1).FooterStyle:id=27,.parent=3"
      _StyleDefs(79)  =   "Splits(1).InactiveStyle:id=28,.parent=5"
      _StyleDefs(80)  =   "Splits(1).SelectedStyle:id=44,.parent=6"
      _StyleDefs(81)  =   "Splits(1).EditorStyle:id=43,.parent=7"
      _StyleDefs(82)  =   "Splits(1).HighlightRowStyle:id=45,.parent=8"
      _StyleDefs(83)  =   "Splits(1).EvenRowStyle:id=46,.parent=9"
      _StyleDefs(84)  =   "Splits(1).OddRowStyle:id=75,.parent=10"
      _StyleDefs(85)  =   "Splits(1).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(86)  =   "Splits(1).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(87)  =   "Splits(1).Columns(0).Style:id=82,.parent=25,.alignment=2,.bgcolor=&HF8ECC9&"
      _StyleDefs(88)  =   ":id=82,.locked=-1"
      _StyleDefs(89)  =   "Splits(1).Columns(0).HeadingStyle:id=79,.parent=26"
      _StyleDefs(90)  =   "Splits(1).Columns(0).FooterStyle:id=80,.parent=27"
      _StyleDefs(91)  =   "Splits(1).Columns(0).EditorStyle:id=81,.parent=43"
      _StyleDefs(92)  =   "Splits(1).Columns(1).Style:id=86,.parent=25,.bgcolor=&HF8ECC9&,.locked=-1"
      _StyleDefs(93)  =   "Splits(1).Columns(1).HeadingStyle:id=83,.parent=26"
      _StyleDefs(94)  =   "Splits(1).Columns(1).FooterStyle:id=84,.parent=27"
      _StyleDefs(95)  =   "Splits(1).Columns(1).EditorStyle:id=85,.parent=43"
      _StyleDefs(96)  =   "Splits(1).Columns(2).Style:id=90,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(97)  =   "Splits(1).Columns(2).HeadingStyle:id=87,.parent=26"
      _StyleDefs(98)  =   "Splits(1).Columns(2).FooterStyle:id=88,.parent=27"
      _StyleDefs(99)  =   "Splits(1).Columns(2).EditorStyle:id=89,.parent=43"
      _StyleDefs(100) =   "Splits(1).Columns(3).Style:id=94,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(101) =   "Splits(1).Columns(3).HeadingStyle:id=91,.parent=26"
      _StyleDefs(102) =   "Splits(1).Columns(3).FooterStyle:id=92,.parent=27"
      _StyleDefs(103) =   "Splits(1).Columns(3).EditorStyle:id=93,.parent=43"
      _StyleDefs(104) =   "Splits(1).Columns(4).Style:id=102,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(105) =   "Splits(1).Columns(4).HeadingStyle:id=99,.parent=26"
      _StyleDefs(106) =   "Splits(1).Columns(4).FooterStyle:id=100,.parent=27"
      _StyleDefs(107) =   "Splits(1).Columns(4).EditorStyle:id=101,.parent=43"
      _StyleDefs(108) =   "Splits(1).Columns(5).Style:id=106,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(109) =   "Splits(1).Columns(5).HeadingStyle:id=103,.parent=26"
      _StyleDefs(110) =   "Splits(1).Columns(5).FooterStyle:id=104,.parent=27"
      _StyleDefs(111) =   "Splits(1).Columns(5).EditorStyle:id=105,.parent=43"
      _StyleDefs(112) =   "Splits(1).Columns(6).Style:id=110,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(113) =   "Splits(1).Columns(6).HeadingStyle:id=107,.parent=26"
      _StyleDefs(114) =   "Splits(1).Columns(6).FooterStyle:id=108,.parent=27"
      _StyleDefs(115) =   "Splits(1).Columns(6).EditorStyle:id=109,.parent=43"
      _StyleDefs(116) =   "Splits(1).Columns(7).Style:id=114,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(117) =   "Splits(1).Columns(7).HeadingStyle:id=111,.parent=26"
      _StyleDefs(118) =   "Splits(1).Columns(7).FooterStyle:id=112,.parent=27"
      _StyleDefs(119) =   "Splits(1).Columns(7).EditorStyle:id=113,.parent=43"
      _StyleDefs(120) =   "Splits(1).Columns(8).Style:id=118,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(121) =   "Splits(1).Columns(8).HeadingStyle:id=115,.parent=26"
      _StyleDefs(122) =   "Splits(1).Columns(8).FooterStyle:id=116,.parent=27"
      _StyleDefs(123) =   "Splits(1).Columns(8).EditorStyle:id=117,.parent=43"
      _StyleDefs(124) =   "Named:id=33:Normal"
      _StyleDefs(125) =   ":id=33,.parent=0"
      _StyleDefs(126) =   "Named:id=34:Heading"
      _StyleDefs(127) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(128) =   ":id=34,.wraptext=-1"
      _StyleDefs(129) =   "Named:id=35:Footing"
      _StyleDefs(130) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(131) =   "Named:id=36:Selected"
      _StyleDefs(132) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(133) =   "Named:id=37:Caption"
      _StyleDefs(134) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(135) =   "Named:id=38:HighlightRow"
      _StyleDefs(136) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(137) =   "Named:id=39:EvenRow"
      _StyleDefs(138) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(139) =   "Named:id=40:OddRow"
      _StyleDefs(140) =   ":id=40,.parent=33"
      _StyleDefs(141) =   "Named:id=41:RecordSelector"
      _StyleDefs(142) =   ":id=41,.parent=34"
      _StyleDefs(143) =   "Named:id=42:FilterBar"
      _StyleDefs(144) =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   3600
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
            Picture         =   "frmPatrimonioNeto.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":220C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":25E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":29C0
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
            Picture         =   "frmPatrimonioNeto.frx":39DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":3B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":3C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":3DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":3F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":409C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":4350
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":44AA
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
            Picture         =   "frmPatrimonioNeto.frx":4604
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":4B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":5138
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":56D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":5C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":6206
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":67A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":6D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatrimonioNeto.frx":72D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   5
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(*) Presione la tecla SUPRIMIR para borrar el campo seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   7185
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   390
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   " Vuelve a cargar los datos almacenados "
      Top             =   450
      Width           =   1665
      Caption         =   " Cargar datos"
      PicturePosition =   327683
      Size            =   "2937;688"
      Picture         =   "frmPatrimonioNeto.frx":786E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdLimpiar 
      Height          =   390
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   " Limpia todas las cuentas contables de la lista "
      Top             =   450
      Width           =   1665
      Caption         =   " Limpiar todo"
      PicturePosition =   327683
      Size            =   "2937;688"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(*) Asigne fórmulas en las cuentas patrimoniales, presionando la tecla F1 en cada celda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Index           =   3
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   7185
   End
End
Attribute VB_Name = "frmManPatrimonioNeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrPatrimonio As New XArrayDB
Dim gsGrupo As String
Dim lArrDetalle(10) As Variant


Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdLimpiar_Click()
    Dim i As Integer, j As Integer
    On Error GoTo serror
    If MsgBox("Desea limpiar todas las cuentas contables de la lista", vbYesNo + vbQuestion) = vbYes Then
        For j = 0 To lArrPatrimonio.Count(1) - 1
            For i = 2 To 8
                lArrPatrimonio(j, i) = ""
            Next i
        Next j
        grdPatrimonio.Refresh
    End If
    
    Exit Sub
serror:
    
End Sub

Private Sub cmdRefresh_Click()
    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    
    GeneraArreglo
    DoEvents
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If tbrOpciones.Buttons(3).Enabled = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                'If respuesta = vbYes Then Call Cancelar
            End If
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        'Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    Call Centrar_form(Me)
    Me.Width = 11025
     
    grdPatrimonio.FetchRowStyle = True
    
    grdPatrimonio.Splits(0).MarqueeStyle = dbgHighlightRow
    grdPatrimonio.HighlightRowStyle = "HighlightRow"
    
    tbrOpciones.Buttons(3).Enabled = True
    GeneraArreglo
    
    SeteaBarraHerramientas tbrOpciones, gsGrupo
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdLimpiar.Enabled = False
        grdPatrimonio.Splits(1).Locked = True
    Else
        cmdLimpiar.Enabled = True
        grdPatrimonio.Splits(1).Locked = False
    End If
    
End Sub

Private Sub GeneraArreglo()
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion

    sql = "spCn_GrabaPatrimonio 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', ''"
    Call GridArreglo(lArrPatrimonio, grdPatrimonio, sql)
    
    
    grdPatrimonio.Splits(1).ScrollBars = dbgBoth
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        grdPatrimonio.Height = Me.Height - 1600
        grdPatrimonio.Width = Me.Width - 300
        tbrOpciones.Width = Me.Width
        grdPatrimonio.ScrollBars = dbgNone
        grdPatrimonio.ScrollBars = dbgAutomatic
        Exit Sub
    End If
    Exit Sub
serror:

End Sub

Private Sub Form_Unload(Cancel As Integer)
     If MsgBox("Desea salir del formulario de Patrimonio Neto", vbYesNo + vbQuestion) = vbNo Then
        Cancel = 1
        Exit Sub
     End If

    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub grdPatrimonio_AfterColEdit(ByVal ColIndex As Integer)

    With grdPatrimonio
        If IsNumeric(.Columns(ColIndex).Value) = False Then
           Mensajes "La cuenta contable debe ser numérica"
           .Columns(ColIndex) = ""
           pSetFocus grdPatrimonio
        Else
           If ExisteCta(.Columns(ColIndex).Value) = "" Then
                Mensajes "La cuenta ingresada no existe"
                .Columns(ColIndex) = ""
                pSetFocus grdPatrimonio
           End If
        End If
    End With

End Sub

Private Sub grdPatrimonio_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    'If lArrPatrimonio(grdPatrimonio.Bookmark, 0) = "00" Then
        Cancel = 1
    '    Exit Sub
    'End If
    
    'If KeyAscii = 46 Then
    '   grdPatrimonio.Columns(ColIndex) = ""
    'End If
    'If IsNumeric(Chr(KeyAscii)) = False Then
    '   KeyAscii = 0
    '   Cancel = 1
    'End If

End Sub

Private Sub grdPatrimonio_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    
    On Error GoTo serror
    'If Split = 1 Then
        If lArrPatrimonio(Bookmark, 0) = "00" Then
            RowStyle.BackColor = &HCEFFF8   'amarillo
        End If
    'End If

    Exit Sub
serror:
End Sub

Public Sub UpdateGrilla()
    On Error Resume Next
    DoEvents
    grdPatrimonio.Update
    DoEvents
End Sub

Private Sub grdPatrimonio_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sCuenta As String
    gsColumna = grdPatrimonio.Col
    
    
    If Mid(gsGrupo, 3, 1) = "1" Or gsGrupo = gsPrivilegioAdmin And gsColumna >= 2 Then
        If KeyCode = vbKeyF1 Then
            
            DoEvents
            sCuenta = CE(grdPatrimonio.Columns(gsColumna))
            DoEvents
           ' If lArrPatrimonio(grdPatrimonio.Bookmark, 0) = "00" Then
                    
                    frmFormulas.pFormula = sCuenta
                    frmFormulas.pObservacion = ""
                    frmFormulas.pFormulario = Me.Name
                    frmFormulas.pTipo = "C"
                    frmFormulas.pMetodo = "PATRIMONIO"
                    frmFormulas.pColumna = gsColumna
                    frmFormulas.Show vbModal
                    pSetFocus frmFormulas.tdbtCuenta
            'Else
            '    On Error Resume Next
            '    sCuenta = Replace(sCuenta, "CTA", "")
            '    LlamaBuscar frmBuscador, "Cuentas", "Cuentas", "CuentasFilt", Me, sCuenta
            'End If
        End If
        
        If KeyCode = vbKeyDelete And gsColumna >= 2 Then
            grdPatrimonio.Columns(gsColumna) = ""
        End If
        
    End If
    
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
    Select Case lControl
           Case "CuentasFilt"
                grdPatrimonio.Columns(grdPatrimonio.Col).Value = "CTA" & param0
           
    End Select
    
End Sub


Private Sub Grabar()
    If ValidaGrabar = True Then
       On Error Resume Next
       grdPatrimonio.Update
       DoEvents
       Mensajes "Las cuentas se grabaron con exito", vbInformation
       
       pSetFocus grdPatrimonio
    End If

    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
End Sub

Private Sub grdPatrimonio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        If lArrPatrimonio(grdPatrimonio.Bookmark, 0) = "00" Then
            'TDBNumLite.BackColor = &HCEFFF8     'amarillo
        Else
            'TDBNumLite.BackColor = gsColorActivado
        End If
        
        gsColumna = grdPatrimonio.Col
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim respuesta As String
    Select Case Button.Index
        Case 3:
                Grabar
                
                
                
        Case 5: 'Editar
        Case 6: Imprimir
        Case 7
                Unload Me
    End Select
End Sub




Private Function ValidaGrabar() As Boolean
    ValidaGrabar = False

    If lArrPatrimonio.Count(1) = 1 And grdPatrimonio.Bookmark = 0 And CE(grdPatrimonio.Columns(0)) = "" Then
        Exit Function
    End If

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas

    grdPatrimonio.Bookmark = grdPatrimonio.Bookmark
    

    Dim lArrDet(2) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPatrimonio", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Function
    End If
    
    For i = 0 To lArrPatrimonio.Count(1) - 1
        If CE(lArrPatrimonio(i, 0)) <> "" Then
            
                If CargaArregloDet(i) = True Then
                    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPatrimonio", lArrDetalle(), False) = False Then
                        Screen.MousePointer = vbNormal
                        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                        Exit Function
                    End If
                End If
            
        End If
    Next

    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal

    Set clsMante = Nothing
    
    Call GeneraArreglo
    
    ValidaGrabar = True
End Function

Private Function CargaArregloDet(item As Integer) As Boolean
    CargaArregloDet = False
    
    lArrDetalle(0) = "INSERTAR"
    lArrDetalle(1) = gsEmpresa
    lArrDetalle(2) = gsAnio
    lArrDetalle(3) = CE(lArrPatrimonio(item, 0))
    lArrDetalle(4) = CE(lArrPatrimonio(item, 2))
    lArrDetalle(5) = CE(lArrPatrimonio(item, 3))
    lArrDetalle(6) = CE(lArrPatrimonio(item, 4))
    lArrDetalle(7) = CE(lArrPatrimonio(item, 5))
    lArrDetalle(8) = CE(lArrPatrimonio(item, 6))
    lArrDetalle(9) = CE(lArrPatrimonio(item, 7))
    lArrDetalle(10) = CE(lArrPatrimonio(item, 8))
    
    Dim cadena As String
    
    cadena = lArrDetalle(4) & lArrDetalle(5) & lArrDetalle(6) & lArrDetalle(7) & _
             lArrDetalle(8) & lArrDetalle(9) & lArrDetalle(10)
    
    If cadena <> "" Then
       CargaArregloDet = True
    End If

End Function

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
    AbreReporteParam gsDSN, Me, rutaReportes & "RptPatrimonioNeto.rpt", crptToWindow, "Reporte de Patrimonio Neto", "", matriz_fecha(), formulas()
    Screen.MousePointer = vbDefault

End Sub

