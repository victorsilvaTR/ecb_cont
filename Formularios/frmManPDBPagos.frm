VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManPDBPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDB - Carga de Forma de Pago"
   ClientHeight    =   6525
   ClientLeft      =   1200
   ClientTop       =   3570
   ClientWidth     =   13035
   Icon            =   "frmManPDBPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   13035
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid70.TDBDropDown tdbdBancos 
      Height          =   3585
      Left            =   5310
      TabIndex        =   16
      Top             =   1890
      Visible         =   0   'False
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   6324
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "Ban_cCodSunat"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "Ban_cNombre"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   0
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid70.TDBDropDown tdbdMedioPago 
      Height          =   2910
      Left            =   225
      TabIndex        =   15
      Top             =   1890
      Visible         =   0   'False
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   5133
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "Tab_cCodSunat"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "Tab_cDescripCampo"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   0
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid70.TDBDropDown tdbdTipoComVen 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   4905
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "Mon_cCodigo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "Mon_cNombreLargo"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   0
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HFFFFFF&"
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
   Begin TDBDate6Ctl.TDBDate tdbdFecha 
      Height          =   300
      Left            =   11565
      TabIndex        =   1
      Tag             =   "enabled"
      Top             =   1755
      Visible         =   0   'False
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   529
      Calendar        =   "frmManPDBPagos.frx":0ECA
      Caption         =   "frmManPDBPagos.frx":0FCC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManPDBPagos.frx":1030
      Keys            =   "frmManPDBPagos.frx":104E
      Spin            =   "frmManPDBPagos.frx":10BA
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
   Begin TDBNumber6Ctl.TDBNumber tdbnNumerico 
      Height          =   240
      Left            =   11430
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   423
      Calculator      =   "frmManPDBPagos.frx":10E2
      Caption         =   "frmManPDBPagos.frx":1102
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManPDBPagos.frx":116E
      Keys            =   "frmManPDBPagos.frx":118C
      Spin            =   "frmManPDBPagos.frx":11D6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,##0.00"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999999
      MinValue        =   -10000
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
   Begin TrueOleDBGrid70.TDBGrid tdbgPDB 
      Height          =   5355
      Left            =   90
      TabIndex        =   3
      Top             =   1035
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9446
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Numero Movimiento"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Numero Voucher"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Codigo Empresa"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Anio"
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
      Columns(6).Caption=   "Tipo PDB"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Item"
      Columns(7).DataField=   "Soles"
      Columns(7).NumberFormat=   "External Editor"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Tipo Compra"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "External Editor"
      Columns(8).DropDown=   "tdbdTipoComVen"
      Columns(8).DropDown.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Tipo Comprobante"
      Columns(9).DataField=   ""
      Columns(9).NumberFormat=   "External Editor"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Fecha Emision/Pago"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "External Editor"
      Columns(10).ExternalEditor=   "tdbdFecha"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Serie Comprob."
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "External Editor"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Numero Comprob."
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "External Editor"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Tipo Persona"
      Columns(13).DataField=   ""
      Columns(13).NumberFormat=   "External Editor"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Tipo doc. Ident."
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "External Editor"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Num. doc. Ident."
      Columns(15).DataField=   ""
      Columns(15).NumberFormat=   "External Editor"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Nombre / Razon Social"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Apellido Paterno"
      Columns(17).DataField=   ""
      Columns(17).NumberFormat=   "External Editor"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Apellido Materno"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Primer Nombre"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Segundo Nombre"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Tipo Moneda"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Cod. Destino"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Num. Destino"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Base Imponible"
      Columns(24).DataField=   ""
      Columns(24).NumberFormat=   "External Editor"
      Columns(24).ExternalEditor=   "tdbnNumerico"
      Columns(24).ExternalEditor.vt=   8
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Monto ISC"
      Columns(25).DataField=   ""
      Columns(25).NumberFormat=   "External Editor"
      Columns(25).ExternalEditor=   "tdbnNumerico"
      Columns(25).ExternalEditor.vt=   8
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Monto IGV"
      Columns(26).DataField=   ""
      Columns(26).NumberFormat=   "External Editor"
      Columns(26).ExternalEditor=   "tdbnNumerico"
      Columns(26).ExternalEditor.vt=   8
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "Monto Otros"
      Columns(27).DataField=   ""
      Columns(27).NumberFormat=   "External Editor"
      Columns(27).ExternalEditor=   "tdbnNumerico"
      Columns(27).ExternalEditor.vt=   8
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Indic. Detracciones"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "Cod. Tasa Detrac."
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "Num. Const. Detrac."
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "Ind. Retenciones"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "Tipo Ref."
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "Serie Ref."
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "Numero Ref."
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "Fecha Ref."
      Columns(35).DataField=   ""
      Columns(35).ExternalEditor=   "tdbdFecha"
      Columns(35).ExternalEditor.vt=   8
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "Base Imp. Ref."
      Columns(36).DataField=   ""
      Columns(36).NumberFormat=   "External Editor"
      Columns(36).ExternalEditor=   "tdbnNumerico"
      Columns(36).ExternalEditor.vt=   8
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "IGV. Ref."
      Columns(37).DataField=   ""
      Columns(37).NumberFormat=   "External Editor"
      Columns(37).ExternalEditor=   "tdbnNumerico"
      Columns(37).ExternalEditor.vt=   8
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "Medio Pago"
      Columns(38).DataField=   ""
      Columns(38).DropDown=   "tdbdMedioPago"
      Columns(38).DropDown.vt=   8
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "Codigo Banco"
      Columns(39).DataField=   ""
      Columns(39).DropDown=   "tdbdBancos"
      Columns(39).DropDown.vt=   8
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "Numero Operacion"
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "Fecha Operacion"
      Columns(41).DataField=   ""
      Columns(41).ExternalEditor=   "tdbdFecha"
      Columns(41).ExternalEditor.vt=   8
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).Caption=   "Monto Operacion"
      Columns(42).DataField=   ""
      Columns(42).NumberFormat=   "External Editor"
      Columns(42).ExternalEditor=   "tdbnNumerico"
      Columns(42).ExternalEditor.vt=   8
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   44
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=44"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1931"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1217"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1138"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1244"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1164"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1164"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1085"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=741"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=661"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=900"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=820"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1111"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1032"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1217"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=512"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(8).AutoDropDown=1"
      Splits(0)._ColumnProps(47)=   "Column(8).DropDownList=1"
      Splits(0)._ColumnProps(48)=   "Column(9).Width=1879"
      Splits(0)._ColumnProps(49)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(9)._WidthInPix=1799"
      Splits(0)._ColumnProps(51)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(53)=   "Column(10).Width=1984"
      Splits(0)._ColumnProps(54)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(10)._WidthInPix=1905"
      Splits(0)._ColumnProps(56)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(57)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(58)=   "Column(11).Width=1799"
      Splits(0)._ColumnProps(59)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(11)._WidthInPix=1720"
      Splits(0)._ColumnProps(61)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(62)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(63)=   "Column(12).Width=2037"
      Splits(0)._ColumnProps(64)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(12)._WidthInPix=1958"
      Splits(0)._ColumnProps(66)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(67)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(68)=   "Column(13).Width=1217"
      Splits(0)._ColumnProps(69)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(13)._WidthInPix=1138"
      Splits(0)._ColumnProps(71)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(72)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(73)=   "Column(14).Width=1402"
      Splits(0)._ColumnProps(74)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(14)._WidthInPix=1323"
      Splits(0)._ColumnProps(76)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(77)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(78)=   "Column(15).Width=2090"
      Splits(0)._ColumnProps(79)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(15)._WidthInPix=2011"
      Splits(0)._ColumnProps(81)=   "Column(15)._ColStyle=514"
      Splits(0)._ColumnProps(82)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(83)=   "Column(16).Width=4763"
      Splits(0)._ColumnProps(84)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(16)._WidthInPix=4683"
      Splits(0)._ColumnProps(86)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(87)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(88)=   "Column(17).Width=2672"
      Splits(0)._ColumnProps(89)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(17)._WidthInPix=2593"
      Splits(0)._ColumnProps(91)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(92)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(93)=   "Column(18).Width=2566"
      Splits(0)._ColumnProps(94)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(18)._WidthInPix=2487"
      Splits(0)._ColumnProps(96)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(97)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(98)=   "Column(19).Width=2619"
      Splits(0)._ColumnProps(99)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(19)._WidthInPix=2540"
      Splits(0)._ColumnProps(101)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(102)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(103)=   "Column(20).Width=2963"
      Splits(0)._ColumnProps(104)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(20)._WidthInPix=2884"
      Splits(0)._ColumnProps(106)=   "Column(20)._ColStyle=8708"
      Splits(0)._ColumnProps(107)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(108)=   "Column(21).Width=1164"
      Splits(0)._ColumnProps(109)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(21)._WidthInPix=1085"
      Splits(0)._ColumnProps(111)=   "Column(21)._ColStyle=8708"
      Splits(0)._ColumnProps(112)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(113)=   "Column(22).Width=1085"
      Splits(0)._ColumnProps(114)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(22)._WidthInPix=1005"
      Splits(0)._ColumnProps(116)=   "Column(22)._ColStyle=8708"
      Splits(0)._ColumnProps(117)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(118)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(119)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(120)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(121)=   "Column(23)._ColStyle=8708"
      Splits(0)._ColumnProps(122)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(123)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(124)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(125)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(126)=   "Column(24)._ColStyle=8706"
      Splits(0)._ColumnProps(127)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(128)=   "Column(25).Width=2302"
      Splits(0)._ColumnProps(129)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(130)=   "Column(25)._WidthInPix=2223"
      Splits(0)._ColumnProps(131)=   "Column(25)._ColStyle=8706"
      Splits(0)._ColumnProps(132)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(133)=   "Column(26).Width=2408"
      Splits(0)._ColumnProps(134)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(26)._WidthInPix=2328"
      Splits(0)._ColumnProps(136)=   "Column(26)._ColStyle=8706"
      Splits(0)._ColumnProps(137)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(138)=   "Column(27).Width=2434"
      Splits(0)._ColumnProps(139)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(140)=   "Column(27)._WidthInPix=2355"
      Splits(0)._ColumnProps(141)=   "Column(27)._ColStyle=514"
      Splits(0)._ColumnProps(142)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(143)=   "Column(28).Width=1826"
      Splits(0)._ColumnProps(144)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(28)._WidthInPix=1746"
      Splits(0)._ColumnProps(146)=   "Column(28)._ColStyle=516"
      Splits(0)._ColumnProps(147)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(148)=   "Column(29).Width=1508"
      Splits(0)._ColumnProps(149)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(29)._WidthInPix=1429"
      Splits(0)._ColumnProps(151)=   "Column(29)._ColStyle=516"
      Splits(0)._ColumnProps(152)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(153)=   "Column(30).Width=3281"
      Splits(0)._ColumnProps(154)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(30)._WidthInPix=3201"
      Splits(0)._ColumnProps(156)=   "Column(30)._ColStyle=516"
      Splits(0)._ColumnProps(157)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(158)=   "Column(31).Width=1455"
      Splits(0)._ColumnProps(159)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(160)=   "Column(31)._WidthInPix=1376"
      Splits(0)._ColumnProps(161)=   "Column(31)._ColStyle=516"
      Splits(0)._ColumnProps(162)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(163)=   "Column(32).Width=926"
      Splits(0)._ColumnProps(164)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(165)=   "Column(32)._WidthInPix=847"
      Splits(0)._ColumnProps(166)=   "Column(32)._ColStyle=516"
      Splits(0)._ColumnProps(167)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(168)=   "Column(33).Width=1535"
      Splits(0)._ColumnProps(169)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(170)=   "Column(33)._WidthInPix=1455"
      Splits(0)._ColumnProps(171)=   "Column(33)._ColStyle=516"
      Splits(0)._ColumnProps(172)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(173)=   "Column(34).Width=1826"
      Splits(0)._ColumnProps(174)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(34)._WidthInPix=1746"
      Splits(0)._ColumnProps(176)=   "Column(34)._ColStyle=516"
      Splits(0)._ColumnProps(177)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(178)=   "Column(35).Width=1852"
      Splits(0)._ColumnProps(179)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(180)=   "Column(35)._WidthInPix=1773"
      Splits(0)._ColumnProps(181)=   "Column(35)._ColStyle=516"
      Splits(0)._ColumnProps(182)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(183)=   "Column(36).Width=2461"
      Splits(0)._ColumnProps(184)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(185)=   "Column(36)._WidthInPix=2381"
      Splits(0)._ColumnProps(186)=   "Column(36)._ColStyle=514"
      Splits(0)._ColumnProps(187)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(188)=   "Column(37).Width=2170"
      Splits(0)._ColumnProps(189)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(190)=   "Column(37)._WidthInPix=2090"
      Splits(0)._ColumnProps(191)=   "Column(37)._ColStyle=514"
      Splits(0)._ColumnProps(192)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(193)=   "Column(38).Width=1535"
      Splits(0)._ColumnProps(194)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(38)._WidthInPix=1455"
      Splits(0)._ColumnProps(196)=   "Column(38)._ColStyle=516"
      Splits(0)._ColumnProps(197)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(198)=   "Column(38).AutoDropDown=1"
      Splits(0)._ColumnProps(199)=   "Column(38).DropDownList=1"
      Splits(0)._ColumnProps(200)=   "Column(39).Width=1667"
      Splits(0)._ColumnProps(201)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(202)=   "Column(39)._WidthInPix=1588"
      Splits(0)._ColumnProps(203)=   "Column(39)._ColStyle=516"
      Splits(0)._ColumnProps(204)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(205)=   "Column(39).AutoDropDown=1"
      Splits(0)._ColumnProps(206)=   "Column(39).DropDownList=1"
      Splits(0)._ColumnProps(207)=   "Column(40).Width=2725"
      Splits(0)._ColumnProps(208)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(209)=   "Column(40)._WidthInPix=2646"
      Splits(0)._ColumnProps(210)=   "Column(40)._ColStyle=516"
      Splits(0)._ColumnProps(211)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(212)=   "Column(41).Width=2196"
      Splits(0)._ColumnProps(213)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(214)=   "Column(41)._WidthInPix=2117"
      Splits(0)._ColumnProps(215)=   "Column(41)._ColStyle=516"
      Splits(0)._ColumnProps(216)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(217)=   "Column(42).Width=2858"
      Splits(0)._ColumnProps(218)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(219)=   "Column(42)._WidthInPix=2778"
      Splits(0)._ColumnProps(220)=   "Column(42)._ColStyle=514"
      Splits(0)._ColumnProps(221)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(222)=   "Column(43).Width=159"
      Splits(0)._ColumnProps(223)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(224)=   "Column(43)._WidthInPix=79"
      Splits(0)._ColumnProps(225)=   "Column(43).AllowSizing=0"
      Splits(0)._ColumnProps(226)=   "Column(43)._ColStyle=8708"
      Splits(0)._ColumnProps(227)=   "Column(43).Order=44"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=32,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=98,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=102,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=114,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=111,.parent=14"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=112,.parent=15"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=113,.parent=17"
      _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=110,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
      _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
      _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
      _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=106,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=14"
      _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=15"
      _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=17"
      _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=118,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
      _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
      _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
      _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=122,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
      _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
      _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
      _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=126,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
      _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
      _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
      _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=130,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
      _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
      _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
      _StyleDefs(133) =   "Splits(0).Columns(24).Style:id=142,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(134) =   ":id=142,.locked=-1"
      _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=139,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=140,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=141,.parent=17"
      _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=138,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(139) =   ":id=138,.locked=-1"
      _StyleDefs(140) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
      _StyleDefs(141) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
      _StyleDefs(142) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
      _StyleDefs(143) =   "Splits(0).Columns(26).Style:id=134,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(144) =   ":id=134,.locked=-1"
      _StyleDefs(145) =   "Splits(0).Columns(26).HeadingStyle:id=131,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(26).FooterStyle:id=132,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(26).EditorStyle:id=133,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(27).Style:id=146,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(149) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(28).Style:id=150,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(153) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
      _StyleDefs(156) =   "Splits(0).Columns(29).Style:id=154,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(157) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=14"
      _StyleDefs(158) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=15"
      _StyleDefs(159) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=17"
      _StyleDefs(160) =   "Splits(0).Columns(30).Style:id=158,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(161) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=14"
      _StyleDefs(162) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=15"
      _StyleDefs(163) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=17"
      _StyleDefs(164) =   "Splits(0).Columns(31).Style:id=162,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(165) =   "Splits(0).Columns(31).HeadingStyle:id=159,.parent=14"
      _StyleDefs(166) =   "Splits(0).Columns(31).FooterStyle:id=160,.parent=15"
      _StyleDefs(167) =   "Splits(0).Columns(31).EditorStyle:id=161,.parent=17"
      _StyleDefs(168) =   "Splits(0).Columns(32).Style:id=166,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(169) =   "Splits(0).Columns(32).HeadingStyle:id=163,.parent=14"
      _StyleDefs(170) =   "Splits(0).Columns(32).FooterStyle:id=164,.parent=15"
      _StyleDefs(171) =   "Splits(0).Columns(32).EditorStyle:id=165,.parent=17"
      _StyleDefs(172) =   "Splits(0).Columns(33).Style:id=170,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(173) =   "Splits(0).Columns(33).HeadingStyle:id=167,.parent=14"
      _StyleDefs(174) =   "Splits(0).Columns(33).FooterStyle:id=168,.parent=15"
      _StyleDefs(175) =   "Splits(0).Columns(33).EditorStyle:id=169,.parent=17"
      _StyleDefs(176) =   "Splits(0).Columns(34).Style:id=174,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(177) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=14"
      _StyleDefs(178) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=15"
      _StyleDefs(179) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=17"
      _StyleDefs(180) =   "Splits(0).Columns(35).Style:id=178,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(181) =   "Splits(0).Columns(35).HeadingStyle:id=175,.parent=14"
      _StyleDefs(182) =   "Splits(0).Columns(35).FooterStyle:id=176,.parent=15"
      _StyleDefs(183) =   "Splits(0).Columns(35).EditorStyle:id=177,.parent=17"
      _StyleDefs(184) =   "Splits(0).Columns(36).Style:id=182,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(185) =   "Splits(0).Columns(36).HeadingStyle:id=179,.parent=14"
      _StyleDefs(186) =   "Splits(0).Columns(36).FooterStyle:id=180,.parent=15"
      _StyleDefs(187) =   "Splits(0).Columns(36).EditorStyle:id=181,.parent=17"
      _StyleDefs(188) =   "Splits(0).Columns(37).Style:id=186,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(189) =   "Splits(0).Columns(37).HeadingStyle:id=183,.parent=14"
      _StyleDefs(190) =   "Splits(0).Columns(37).FooterStyle:id=184,.parent=15"
      _StyleDefs(191) =   "Splits(0).Columns(37).EditorStyle:id=185,.parent=17"
      _StyleDefs(192) =   "Splits(0).Columns(38).Style:id=210,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(193) =   "Splits(0).Columns(38).HeadingStyle:id=207,.parent=14"
      _StyleDefs(194) =   "Splits(0).Columns(38).FooterStyle:id=208,.parent=15"
      _StyleDefs(195) =   "Splits(0).Columns(38).EditorStyle:id=209,.parent=17"
      _StyleDefs(196) =   "Splits(0).Columns(39).Style:id=206,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(197) =   "Splits(0).Columns(39).HeadingStyle:id=203,.parent=14"
      _StyleDefs(198) =   "Splits(0).Columns(39).FooterStyle:id=204,.parent=15"
      _StyleDefs(199) =   "Splits(0).Columns(39).EditorStyle:id=205,.parent=17"
      _StyleDefs(200) =   "Splits(0).Columns(40).Style:id=202,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(201) =   "Splits(0).Columns(40).HeadingStyle:id=199,.parent=14"
      _StyleDefs(202) =   "Splits(0).Columns(40).FooterStyle:id=200,.parent=15"
      _StyleDefs(203) =   "Splits(0).Columns(40).EditorStyle:id=201,.parent=17"
      _StyleDefs(204) =   "Splits(0).Columns(41).Style:id=198,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(205) =   "Splits(0).Columns(41).HeadingStyle:id=195,.parent=14"
      _StyleDefs(206) =   "Splits(0).Columns(41).FooterStyle:id=196,.parent=15"
      _StyleDefs(207) =   "Splits(0).Columns(41).EditorStyle:id=197,.parent=17"
      _StyleDefs(208) =   "Splits(0).Columns(42).Style:id=194,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(209) =   "Splits(0).Columns(42).HeadingStyle:id=191,.parent=14"
      _StyleDefs(210) =   "Splits(0).Columns(42).FooterStyle:id=192,.parent=15"
      _StyleDefs(211) =   "Splits(0).Columns(42).EditorStyle:id=193,.parent=17"
      _StyleDefs(212) =   "Splits(0).Columns(43).Style:id=190,.parent=13,.locked=-1"
      _StyleDefs(213) =   "Splits(0).Columns(43).HeadingStyle:id=187,.parent=14"
      _StyleDefs(214) =   "Splits(0).Columns(43).FooterStyle:id=188,.parent=15"
      _StyleDefs(215) =   "Splits(0).Columns(43).EditorStyle:id=189,.parent=17"
      _StyleDefs(216) =   "Named:id=33:Normal"
      _StyleDefs(217) =   ":id=33,.parent=0"
      _StyleDefs(218) =   "Named:id=34:Heading"
      _StyleDefs(219) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(220) =   ":id=34,.wraptext=-1"
      _StyleDefs(221) =   "Named:id=35:Footing"
      _StyleDefs(222) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(223) =   "Named:id=36:Selected"
      _StyleDefs(224) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(225) =   "Named:id=37:Caption"
      _StyleDefs(226) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(227) =   "Named:id=38:HighlightRow"
      _StyleDefs(228) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(229) =   "Named:id=39:EvenRow"
      _StyleDefs(230) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(231) =   "Named:id=40:OddRow"
      _StyleDefs(232) =   ":id=40,.parent=33"
      _StyleDefs(233) =   "Named:id=41:RecordSelector"
      _StyleDefs(234) =   ":id=41,.parent=34"
      _StyleDefs(235) =   "Named:id=42:FilterBar"
      _StyleDefs(236) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   225
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
      _PropDict       =   $"frmManPDBPagos.frx":11FE
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
      Left            =   900
      TabIndex        =   5
      Tag             =   "enabled"
      Top             =   585
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
      _PropDict       =   $"frmManPDBPagos.frx":1285
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
   Begin MSForms.CommandButton cmdPreliminar 
      Height          =   375
      Left            =   9450
      TabIndex        =   17
      Top             =   585
      Width           =   1575
      Caption         =   " Vista Preliminar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":130C
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
      Left            =   180
      TabIndex        =   14
      Top             =   270
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
      Left            =   180
      TabIndex        =   13
      Top             =   630
      Width           =   420
   End
   Begin MSForms.CommandButton cmdListar 
      Height          =   375
      Left            =   4455
      TabIndex        =   12
      ToolTipText     =   "Cargar nueva Configuración"
      Top             =   180
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":18A6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   180
      Width           =   1575
      Caption         =   " Insertar Item"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":1E40
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   7785
      TabIndex        =   10
      ToolTipText     =   "Eliminar el movimientos seleccionado"
      Top             =   180
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
      Left            =   7785
      TabIndex        =   9
      ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
      Top             =   585
      Width           =   1575
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":23DA
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   4455
      TabIndex        =   8
      ToolTipText     =   "Grabar modificaciones"
      Top             =   585
      Width           =   1575
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":2974
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdTodos 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      ToolTipText     =   "Insertar todos los movimientos del libro y mes seleccionado"
      Top             =   585
      Width           =   1575
      Caption         =   " Insertar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":2F0E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   9450
      TabIndex        =   6
      Top             =   180
      Width           =   1575
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManPDBPagos.frx":34A8
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmManPDBPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrDatos As New XArrayDB
Dim lrsTipoConVen As New ADODB.Recordset
Dim lrsMedioPago As New ADODB.Recordset
Dim lrsBancos As New ADODB.Recordset
Dim lControl As String
Dim lArrDet() As Variant
Dim sw As Boolean
Dim gsGrupo As String
Dim cTipoPDB As String
Dim NUM_FILAS As Integer
Dim NUM_COLUMNAS As Integer

Dim nCol_cNummov As Integer
Dim nCol_cNumVoucher As Integer
Dim nCol_cEmpresa As Integer
Dim nCol_cPanAnio As Integer
Dim nCol_cPeriodo As Integer
Dim nCol_cTipoLibro As Integer
Dim nCol_cTipoPDB As Integer
Dim nCol_cItem As Integer
Dim nCol_cTipo As Integer
Dim nCol_cTipoComp As Integer
Dim nCol_dFechaComp As Integer
Dim nCol_cSerieComp As Integer
Dim nCol_cNumComp As Integer
Dim nCol_cTipoPer As Integer
Dim nCol_cTipoDocPer As Integer
Dim nCol_cNumDocPer As Integer
Dim nCol_cRazon As Integer
Dim nCol_cAppPat As Integer
Dim nCol_cAppMat As Integer
Dim nCol_cPriNom As Integer
Dim nCol_cSegNom As Integer
Dim nCol_cTipoMon As Integer
Dim nCol_cCodDestino As Integer
Dim nCol_cNumDestino As Integer
Dim nCol_nBaseImp As Integer
Dim nCol_nISC As Integer
Dim nCol_nIGV As Integer
Dim nCol_nOtros As Integer
Dim nCol_cIndDetra As Integer
Dim nCol_cTasaDetra As Integer
Dim nCol_cConsDetra As Integer
Dim nCol_cIndReten As Integer
Dim nCol_cTipoRef As Integer
Dim nCol_cSerieRef As Integer
Dim nCol_cNumeroRef As Integer
Dim nCol_cFechaRef As Integer
Dim nCol_nBaseImpRef As Integer
Dim nCol_nIGVRef As Integer
Dim nCol_cMedio As Integer
Dim nCol_cCodBan As Integer
Dim nCol_cNumOp As Integer
Dim nCol_cFecOp As Integer
Dim nCol_nMonOp As Integer

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub EliminaTodo()
        Dim clsMante As New clsMantoTablas
        
        Call EliminaArreglo
        
        clsMante.InicializaClase
        clsMante.BeginTrans
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoPDB", lArrDet(), False) = False Then
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
        
End Sub

Private Sub cmdEliminarTodo_Click()
    If ValidaCampos = False Then Exit Sub
    
    cmdEliminarTodo.Enabled = False
    DoEvents

    If MsgBox("Deseas eliminar todos los registros de importación", vbYesNo + vbInformation) = vbYes Then
        EliminaTodo
        llenaGrilla
    End If
    
    cmdEliminarTodo.Enabled = True
    
End Sub

Private Sub cmdEliminaItem_Click()
    If ValidaCampos = False Then Exit Sub
    
    If CE(tdbgPDB.Columns(tdbgPDB.Bookmark)) <> "" Then
        cmdEliminaItem.Enabled = False
        DoEvents
        
        lArrDatos.DeleteRows (tdbgPDB.Bookmark)

        Set tdbgPDB.Array = lArrDatos
        Call UpdateGrilla(tdbgPDB)
        Call RebindGrilla(tdbgPDB)
        
        
        DoEvents
        cmdEliminaItem.Enabled = True
        
    End If
    
    Call UpdateGrilla(tdbgPDB)
    Call RefreshGrilla(tdbgPDB)
End Sub

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



Public Function Grabar() As Boolean
    
    If ValidaPDBReglas = False Then Call RefreshGrilla(tdbgPDB):       Exit Function
    
    Dim i As Integer
    Dim clsMante As New clsMantoTablas
    
    
    
    Grabar = True
    Dim Fila As Integer

    Fila = CuentaFilas
    '--------------------------------------------
    DoEvents
    Call EliminaTodo
    DoEvents
    '--------------------------------------------
    Screen.MousePointer = vbHourglass
    clsMante.InicializaClase
    clsMante.BeginTrans
    
    On Error GoTo ERROR
        For i = 0 To Fila - 1
            If CE(lArrDatos(i, 1)) <> "" Then
                Call CargaArregloDet(i)
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoPDB", lArrDet(), False) = False Then
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
    Set clsMante = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    Exit Function
ERROR:
    Grabar = False
    Screen.MousePointer = vbNormal
End Function


Private Sub cmdGrabar_Click()
    If ValidaCampos = False Then Exit Sub
    
    cmdGrabar.Enabled = False
    DoEvents
    If Grabar = True Then
        Mensajes "Datos se grabaron con exito.", vbInformation
        
    End If
    DoEvents
    cmdGrabar.Enabled = True
End Sub

Private Sub EliminaArreglo()
    ReDim lArrDet(7)
    lArrDet(0) = "ELIMINAR_TODOS"
    lArrDet(1) = ""
    lArrDet(2) = ""
    lArrDet(3) = gsEmpresa
    lArrDet(4) = gsAnio
    lArrDet(5) = tdbcMes.BoundText
    lArrDet(6) = tdbcLibro.BoundText
    lArrDet(7) = cTipoPDB
        
End Sub

Private Sub CargaArregloDet(item As Integer)
    On Error Resume Next
    Dim i As Integer
    i = 0
    ReDim lArrDet(44) As Variant
    lArrDet(0) = "INSERTAR"
    lArrDet(1) = CE(lArrDatos(item, 0 + i))
    lArrDet(2) = CE(lArrDatos(item, 1 + i))
    lArrDet(3) = CE(lArrDatos(item, 2 + i))
    lArrDet(4) = CE(lArrDatos(item, 3 + i))
    lArrDet(5) = CE(lArrDatos(item, 4 + i))
    lArrDet(6) = CE(lArrDatos(item, 5 + i))
    lArrDet(7) = CE(lArrDatos(item, 6 + i))
    lArrDet(8) = CE(lArrDatos(item, 7 + i))
    lArrDet(9) = CE(lArrDatos(item, 8 + i))
    
    lArrDet(10) = CE(lArrDatos(item, 9 + i))
    lArrDet(11) = CE(lArrDatos(item, 10 + i))
    lArrDet(12) = CE(lArrDatos(item, 11 + i))
    lArrDet(13) = CE(lArrDatos(item, 12 + i))
    lArrDet(14) = CE(lArrDatos(item, 13 + i))
    lArrDet(15) = CE(lArrDatos(item, 14 + i))
    lArrDet(16) = CE(lArrDatos(item, 15 + i))
    lArrDet(17) = Left(CE(lArrDatos(item, 16 + i)), 40)
    
    lArrDet(18) = CE(lArrDatos(item, 17 + i))
    lArrDet(19) = CE(lArrDatos(item, 18 + i))
    
    lArrDet(20) = CE(lArrDatos(item, 19 + i))
    lArrDet(21) = CE(lArrDatos(item, 20 + i))
    lArrDet(22) = CE(lArrDatos(item, 21 + i))
    lArrDet(23) = CE(lArrDatos(item, 22 + i))
    lArrDet(24) = NE(lArrDatos(item, 23 + i))
    lArrDet(25) = NE(lArrDatos(item, 24 + i))
    lArrDet(26) = NE(lArrDatos(item, 25 + i))
    lArrDet(27) = NE(lArrDatos(item, 26 + i))
    
    lArrDet(28) = CE(lArrDatos(item, 27 + i))
    lArrDet(29) = CE(lArrDatos(item, 28 + i))
    
    lArrDet(30) = CE(lArrDatos(item, 29 + i))
    lArrDet(31) = CE(lArrDatos(item, 30 + i))
    lArrDet(32) = CE(lArrDatos(item, 31 + i))
    lArrDet(33) = CE(lArrDatos(item, 32 + i))
    lArrDet(34) = CE(lArrDatos(item, 33 + i))
    lArrDet(35) = CE(lArrDatos(item, 34 + i))
    lArrDet(36) = CE(lArrDatos(item, 35 + i))
    lArrDet(37) = NE(lArrDatos(item, 36 + i))
    lArrDet(38) = NE(lArrDatos(item, 37 + i))
    lArrDet(39) = CE(lArrDatos(item, 38 + i))
    lArrDet(40) = CE(lArrDatos(item, 39 + i))

    lArrDet(41) = CE(lArrDatos(item, 40 + i))
    lArrDet(42) = FE(lArrDatos(item, 41 + i))
    lArrDet(43) = CE(lArrDatos(item, 42 + i))
    lArrDet(44) = gsUsuario
    
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False

    If tdbcMes.BoundText = "" Then
        Mensajes "Seleccione un periodo de la lista"
        Exit Function
    End If
    
    If tdbcLibro.BoundText = "" Then
        Mensajes "Seleccione un libro de la lista"
        Exit Function
    End If
    
    
    ValidaCampos = True
End Function

Private Sub cmdInsertarItem_Click()
    
    If ValidaCampos = False Then Exit Sub
    
    cmdInsertarItem.Enabled = False
    DoEvents
    gsPeriodoCOA = Me.tdbcMes.BoundText
    Call LlamaBuscar(frmBusCoa, "Provisiones", lControl, "Provisiones", Me, tdbcMes.BoundText, Me.tdbcLibro.BoundText)
    DoEvents
    cmdInsertarItem.Enabled = True
End Sub

Private Sub cmdListar_Click()
    If ValidaCampos = False Then Exit Sub
    
    cmdListar.Enabled = False
    DoEvents
    Call llenaGrilla
    DoEvents
    cmdListar.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Function BuscaCelda() As Integer
    Dim i  As Integer
    BuscaCelda = 0
    For i = 0 To lArrDatos.Count(1) - 1
        If CE(lArrDatos(i, 2)) = "" Then
            BuscaCelda = i
            Exit For
        End If
    Next i
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdTodos_Click()
    If ValidaCampos = False Then Exit Sub
    
    If Grabar = False Then Exit Sub
    
    ' *** Jalar todos los datos dependiendo del Tipo de Libro
    cmdTodos.Enabled = False
    DoEvents
    
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim i As Integer
    Dim Fila As Integer
    Dim lrsProvision As New ADODB.Recordset
    Dim cTipoPersona As String
    Dim cMoneda As String
    Dim cBaseImp  As String

    Dim nImporte As Double
    
    Set clDatos = New clsMantoTablas
    Set lrsProvision = New ADODB.Recordset
    sqlSp = "spCn_GrabaAsientoPDB 'BUSCA_PROVISIONES', '','','" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcLibro.BoundText & "','" & cTipoPDB & "' "
    arrDatos = Array(sqlSp)
    Set lrsProvision = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsProvision.State <> 0 Then
        ' *** Cargar los datos de la grilla
        Screen.MousePointer = vbHourglass
        
        i = 0
        Fila = BuscaCelda
        
        Do While Not lrsProvision.EOF
            lArrDatos(Fila, 0 + i) = CE(lrsProvision!Ase_cNummov) 'nummov
            lArrDatos(Fila, 1 + i) = CE(lrsProvision!Ase_nVoucher) 'voucher
            lArrDatos(Fila, 2 + i) = CE(lrsProvision!Emp_cCodigo) 'empresa
            lArrDatos(Fila, 3 + i) = CE(lrsProvision!Pan_cAnio) 'anio
            lArrDatos(Fila, 4 + i) = CE(lrsProvision!Per_cPeriodo) 'periodo
            lArrDatos(Fila, 5 + i) = CE(lrsProvision!Lib_cTipoLibro) 'libro
            lArrDatos(Fila, 6 + i) = cTipoPDB  'tipo pdb
            lArrDatos(Fila, 7 + i) = CE(lrsProvision!asd_nitem) 'item
            
            lArrDatos(Fila, 8 + i) = "01" 'com/ven
            lArrDatos(Fila, 9 + i) = CE(lrsProvision!Asd_cTipoDoc) 'tipo doc
            lArrDatos(Fila, 10 + i) = FE(lrsProvision!Asd_dFecDoc) 'fecha doc
            
            If CE(lrsProvision!Asd_cTipoDoc) = "12" Then
                lArrDatos(Fila, 11 + i) = ""
            Else
                lArrDatos(Fila, 11 + i) = CE(lrsProvision!Asd_cSerieDoc) 'serie doc
            End If
            

            lArrDatos(Fila, 12 + i) = CE(lrsProvision!Asd_cNumDoc) 'numero doc
            
            
            cTipoPersona = CE(lrsProvision!Ent_cFlagPersona) 'tipo persona
            
            Select Case cTipoPersona
                Case "N": lArrDatos(Fila, 13 + i) = "01" 'tipo doc ident
                Case "J": lArrDatos(Fila, 13 + i) = "02" 'tipo doc ident
                Case Else: lArrDatos(Fila, 13 + i) = "03" 'tipo doc ident
            End Select
            
            lArrDatos(Fila, 14 + i) = CE(lrsProvision!Ent_cTipoDoc)   'tipo doc ident
            lArrDatos(Fila, 15 + i) = CE(lrsProvision!Ent_nRuc)  'numdoc ident
            lArrDatos(Fila, 16 + i) = "" 'nombre o razon social
            lArrDatos(Fila, 17 + i) = "" 'apellido paterno
            lArrDatos(Fila, 18 + i) = "" 'apellido materno
            lArrDatos(Fila, 19 + i) = "" 'primer nombre
            
            lArrDatos(Fila, 20 + i) = "" 'segundo nombre
            
            cMoneda = CE(lrsProvision!asd_ctipomoneda)    'codigo moneda
            
            lArrDatos(Fila, 21 + i) = "" 'codigo moneda
            lArrDatos(Fila, 22 + i) = "" 'codigo de destino
            lArrDatos(Fila, 23 + i) = "" 'numero de destino
            
            Select Case cMoneda
                Case gsMonedaNac: nImporte = NE(lrsProvision!Soles)
                Case Else: nImporte = NE(lrsProvision!Dolares)
            End Select
            
            lArrDatos(Fila, 24 + i) = Redondear(nImporte / 1.19, 2)  ' nBaseImp
            
            If CE(lrsProvision!Asd_cTipoDoc) = "05" Then 'tipo doc
                lArrDatos(Fila, 25 + i) = "" 'nISC
            Else
                lArrDatos(Fila, 25 + i) = 0 'nISC
            End If
            

            lArrDatos(Fila, 26 + i) = Redondear((0.19 / 1.19) * nImporte, 2)   'nIGV
            lArrDatos(Fila, 27 + i) = 0 'nOtros
            
            
            lArrDatos(Fila, 28 + i) = "" 'indicador de detracciones
            lArrDatos(Fila, 29 + i) = ""  'codigo de tasa detracciones
            lArrDatos(Fila, 30 + i) = "" 'numero constancia detracciones
            lArrDatos(Fila, 31 + i) = ""  'indicador de retenciones
            lArrDatos(Fila, 32 + i) = ""  'ref tipo doc
            lArrDatos(Fila, 33 + i) = "" 'ref serie doc
            lArrDatos(Fila, 34 + i) = ""  'ref num doc
            lArrDatos(Fila, 35 + i) = "" 'ref fecha doc
            lArrDatos(Fila, 36 + i) = 0 'ref base imp
            lArrDatos(Fila, 37 + i) = 0 'reg igv
            
            lArrDatos(Fila, 38 + i) = CE(lrsProvision!asd_cformapago) 'medio de pago
            lArrDatos(Fila, 39 + i) = "" 'codigo banco
            lArrDatos(Fila, 40 + i) = "" 'num operacion
            lArrDatos(Fila, 41 + i) = "" 'fecha op
            lArrDatos(Fila, 42 + i) = nImporte 'monto op
            
            lArrDatos.AppendRows
            lrsProvision.MoveNext
            
            Fila = Fila + 1
        Loop
        Screen.MousePointer = vbDefault

        'If Grabar = True Then
        '    Mensajes "Los datos se insertaron correctamente", vbInformation
        ' Else
        '    Mensajes "No se pudo insertar las importaciones.", vbInformation + vbOKOnly
        ' End If
          
    Else
        Mensajes "No se encontraron movimientos para el mes y libro seleccionado", vbInformation
    End If
    
    Call CerrarRecordSet(lrsProvision)

    Call RebindGrilla(tdbgPDB)
    Call RefreshGrilla(tdbgPDB)


    On Error Resume Next
    If i >= 0 Then tdbgPDB.Bookmark = 0
    cmdTodos.Enabled = True
    
End Sub

Private Sub IniciaVariables()
 nCol_cNummov = 0
 nCol_cNumVoucher = 1
 nCol_cEmpresa = 2
 nCol_cPanAnio = 3
 nCol_cPeriodo = 4
 nCol_cTipoLibro = 5
 nCol_cTipoPDB = 6
 nCol_cItem = 7
 nCol_cTipo = 8
 nCol_cTipoComp = 9
 nCol_dFechaComp = 10
 nCol_cSerieComp = 11
 nCol_cNumComp = 12
 
 nCol_cTipoPer = 13
 nCol_cTipoDocPer = 14
 nCol_cNumDocPer = 15
 
 nCol_cRazon = 16
 nCol_cAppPat = 17
 nCol_cAppMat = 18
 nCol_cPriNom = 19
 nCol_cSegNom = 20
 nCol_cTipoMon = 21
 nCol_cCodDestino = 22
 nCol_cNumDestino = 23
 nCol_nBaseImp = 24
 nCol_nISC = 25
 nCol_nIGV = 26
 nCol_nOtros = 27
 nCol_cIndDetra = 28
 nCol_cTasaDetra = 29
 nCol_cConsDetra = 30
 nCol_cIndReten = 31
 nCol_cTipoRef = 32
 nCol_cSerieRef = 33
 nCol_cNumeroRef = 34
 nCol_cFechaRef = 35
 nCol_nBaseImpRef = 36
 nCol_nIGVRef = 37
 nCol_cMedio = 38
 nCol_cCodBan = 39
 nCol_cNumOp = 40
 nCol_cFecOp = 41
 nCol_nMonOp = 42
End Sub

Private Sub CargaCombos()
    Dim sqlcombos  As String
    
    Call CerrarRecordSet(lrsTipoConVen)
    lrsTipoConVen.CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
    lrsTipoConVen.CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
    lrsTipoConVen.LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
    lrsTipoConVen.Fields.Append "CODIGO", adChar, 2
    lrsTipoConVen.Fields.Append "DESCRIPCION", adVarChar, 20
    
    lrsTipoConVen.Open
    
    lrsTipoConVen.AddNew
    lrsTipoConVen.Fields("CODIGO") = "01"
    lrsTipoConVen.Fields("DESCRIPCION") = "COMPRA INTERNA"
    lrsTipoConVen.AddNew
    lrsTipoConVen.Fields("CODIGO") = "02"
    lrsTipoConVen.Fields("DESCRIPCION") = "COMPRA EXTERNA"
    lrsTipoConVen.Update
    
    Set tdbdTipoComVen.DataSource = lrsTipoConVen
    tdbdTipoComVen.Columns(0).DataField = "CODIGO"
    tdbdTipoComVen.Columns(1).DataField = "DESCRIPCION"
    
    '-------------------------
    Call CerrarRecordSet(lrsMedioPago)
    
    Set tdbdMedioPago.DataSource = Nothing

    sqlcombos = "select * from tabla where tab_ctabla='074' and emp_ccodigo='" & gsEmpresa & "'"
    Set lrsMedioPago = fRetornaRS(sqlcombos)
    Set tdbdMedioPago.DataSource = lrsMedioPago

    '-------------------------
    Call CerrarRecordSet(lrsBancos)
    
    Set tdbdBancos.DataSource = Nothing

    sqlcombos = "spCNT_MANTBANCO 'SEL_ALL', '" & gsEmpresa & "'"
    Set lrsBancos = fRetornaRS(sqlcombos)
    Set tdbdBancos.DataSource = lrsBancos

    
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    Call Centrar_form(Me)
    
    Call IniciaVariables
    Call CargaCombos
    
    cTipoPDB = "P"
    
    NUM_FILAS = 0
    NUM_COLUMNAS = 51

    Dim sqlcombos As String
    
    
    pCargaCfgLibro
    sw = False
    
    Call LlenaComboMesAddItem(tdbcMes, True, True, "[ Seleccione Mes]")
    
    
    Dim registros As Integer
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' and LIB_CTIPOLIBRO='" & lsLibroCajEgr & "' " & _
                "ORDER BY LIB_CDESCRIPCION "
    
    registros = LlenarComboAddItem(tdbcLibro, sqlcombos, True, True, "[ Seleccione Libro ]")
    
    
    If registros > 0 Then
        DoEvents
        
        Call llenaGrilla
    
        tdbcLibro.Enabled = True
    
    Else
        Mensajes "No se crearon los libros contables en el sistema, ingreselos en mantenimiento de libros", vbOKOnly + vbInformation
        DesactivaBotones False
    End If
    
    Call ConfigurarColumnas
    
    lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
    
    gsGrupo = "0000"
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        DesactivaBotones False
        tdbgPDB.Splits(0).Locked = True
    Else
        DesactivaBotones True
        tdbgPDB.Splits(0).Locked = False
    End If
    
End Sub

Private Sub ConfigurarColumnas()
    Call OcultaColumna(nCol_cNummov)
    Call OcultaColumna(nCol_cEmpresa)
    Call OcultaColumna(nCol_cPanAnio)
    Call OcultaColumna(nCol_cPeriodo)
    Call OcultaColumna(nCol_cTipoLibro)
    Call OcultaColumna(nCol_cTipoPDB)
    Call OcultaColumna(nCol_cItem)
    Call OcultaColumna(nCol_cIndDetra)
    Call OcultaColumna(nCol_cIndReten)
    'Call OcultaColumna(nCol_cTipoMon)
    'Call OcultaColumna(nCol_dFechaComp)
    Call OcultaColumna(nCol_cRazon)
    Call OcultaColumna(nCol_cAppPat)
    Call OcultaColumna(nCol_cPriNom)
    Call OcultaColumna(nCol_cAppMat)
    Call OcultaColumna(nCol_cSegNom)
    Call OcultaColumna(nCol_cCodDestino)
    Call OcultaColumna(nCol_cNumDestino)
    
    Call OcultaColumna(nCol_nISC)
    
    Call OcultaColumna(nCol_nOtros)
    Call OcultaColumna(nCol_cTasaDetra)
    Call OcultaColumna(nCol_cConsDetra)
    Call OcultaColumna(nCol_cTipoRef)
    Call OcultaColumna(nCol_cSerieRef)
    Call OcultaColumna(nCol_cNumeroRef)
    Call OcultaColumna(nCol_cFechaRef)
    Call OcultaColumna(nCol_nBaseImpRef)
    Call OcultaColumna(nCol_nIGVRef)
    
    Call BloqueaColumna(nCol_nBaseImp)
    Call BloqueaColumna(nCol_nIGV)
    Call BloqueaColumna(nCol_cNumVoucher)
    Call BloqueaColumna(nCol_cTipoComp)
    Call BloqueaColumna(nCol_dFechaComp)
    Call BloqueaColumna(nCol_cSerieComp)
    Call BloqueaColumna(nCol_cNumComp)
    Call BloqueaColumna(nCol_cTipoPer)
    Call BloqueaColumna(nCol_cTipoDocPer)
    Call BloqueaColumna(nCol_cNumDocPer)
    Call BloqueaColumna(nCol_cRazon)
    Call BloqueaColumna(nCol_cTipoMon)
    Call BloqueaColumna(nCol_cCodDestino)
    Call BloqueaColumna(nCol_cTipoRef)
    Call BloqueaColumna(nCol_cSerieRef)
    Call BloqueaColumna(nCol_cNumeroRef)
    Call BloqueaColumna(nCol_cFechaRef)
    
End Sub

Private Sub OcultaColumna(nCol As Integer)
    On Error GoTo serror
    With tdbgPDB
        .Columns(nCol).Visible = False
        .Columns(nCol).Width = 0
        .Splits(0).Columns(nCol).AllowFocus = False
        .Splits(0).Columns(nCol).AllowSizing = False
    End With
    Exit Sub
serror:
End Sub

Private Sub BloqueaColumna(nCol As Integer)
    On Error GoTo serror
    With tdbgPDB
        .Columns(nCol).BackColor = gsColorDesactivado
        .Splits(0).Columns(nCol).AllowFocus = False
        .Splits(0).Columns(nCol).AllowSizing = True
    End With
    Exit Sub
serror:
End Sub

Private Sub DesactivaBotones(Valor As Boolean)
    Me.cmdEliminaItem.Enabled = Valor
    Me.cmdEliminarTodo.Enabled = Valor
    Me.cmdGrabar.Enabled = Valor
    Me.cmdInsertarItem.Enabled = Valor
    'Me.cmdListar.Enabled = Valor
    Me.cmdTodos.Enabled = Valor
    
    DoEvents
End Sub



Public Sub llenaGrilla()
    'If ValidaCampos = False Then Exit Sub
    
    Dim sqlcombos As String
    Dim rsArreglo As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    Dim i As Integer
    Dim Col As Integer
    
    If tdbcLibro.BoundText = "" Then
        'Mensajes "Seleccione primero el libro"
        pSetFocus tdbcLibro
        Exit Sub
    End If
    
    sqlcombos = "spCn_GrabaAsientoPDB 'BUSCARTODOS', '','', '" & gsEmpresa & "',  '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & tdbcLibro.BoundText & "','" & cTipoPDB & "' "

    arrDatos = Array(sqlcombos)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo Is Nothing Then
        Screen.MousePointer = vbNormal
        Set rsArreglo = Nothing
        lArrDatos.Clear
        lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
        
        Set tdbgPDB.Array = lArrDatos
        Call UpdateGrilla(tdbgPDB)
        Call RebindGrilla(tdbgPDB)
        
        Exit Sub
    End If
    
    lArrDatos.Clear
    
    lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
    
    If lArrDatos.Count(1) = 0 And lArrDatos.Count(2) = 0 Then
        lArrDatos.AppendRows
        Call RebindGrilla(tdbgPDB)
        Call UpdateGrilla(tdbgPDB)
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
        lArrDatos(i, 16 + Col) = CE(rsArreglo(15).Value)
        lArrDatos(i, 17 + Col) = CE(rsArreglo(16).Value)
        lArrDatos(i, 18 + Col) = CE(rsArreglo(17).Value)
        lArrDatos(i, 19 + Col) = CE(rsArreglo(18).Value)
        lArrDatos(i, 20 + Col) = CE(rsArreglo(19).Value)
        lArrDatos(i, 21 + Col) = CE(rsArreglo(20).Value)
        lArrDatos(i, 22 + Col) = CE(rsArreglo(21).Value)
        lArrDatos(i, 23 + Col) = CE(rsArreglo(22).Value)
        lArrDatos(i, 24 + Col) = CE(rsArreglo(23).Value)
        lArrDatos(i, 25 + Col) = CE(rsArreglo(24).Value)
        lArrDatos(i, 26 + Col) = CE(rsArreglo(25).Value)
        lArrDatos(i, 27 + Col) = CE(rsArreglo(26).Value)
        lArrDatos(i, 28 + Col) = CE(rsArreglo(27).Value)
        lArrDatos(i, 29 + Col) = CE(rsArreglo(28).Value)
        lArrDatos(i, 30 + Col) = CE(rsArreglo(29).Value)
        lArrDatos(i, 31 + Col) = CE(rsArreglo(30).Value)
        lArrDatos(i, 32 + Col) = CE(rsArreglo(31).Value)
        lArrDatos(i, 33 + Col) = CE(rsArreglo(32).Value)
        lArrDatos(i, 34 + Col) = CE(rsArreglo(33).Value)
        lArrDatos(i, 35 + Col) = CE(rsArreglo(34).Value)
        lArrDatos(i, 36 + Col) = CE(rsArreglo(35).Value)
        lArrDatos(i, 37 + Col) = CE(rsArreglo(36).Value)
        lArrDatos(i, 38 + Col) = CE(rsArreglo(37).Value)
        lArrDatos(i, 39 + Col) = CE(rsArreglo(38).Value)
        lArrDatos(i, 40 + Col) = CE(rsArreglo(39).Value)
        lArrDatos(i, 41 + Col) = CE(rsArreglo(40).Value)
        lArrDatos(i, 42 + Col) = CE(rsArreglo(41).Value)
        lArrDatos(i, 43 + Col) = CE(rsArreglo(42).Value)
        
        lArrDatos.AppendRows
        
        rsArreglo.MoveNext
        
        i = i + 1
    Loop
    
    Set tdbgPDB.Array = lArrDatos
    Call RebindGrilla(tdbgPDB)
    Call UpdateGrilla(tdbgPDB)
    
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Dim psql$
    Select Case lControl
            Case "Provisiones"
                If ValidaMovimiento(frmBusCoa.tdbgProvisiones.Columns(20), frmBusCoa.tdbgProvisiones.Columns(5), frmBusCoa.tdbgProvisiones.Columns(21)) = True Then
                    Call LlenaProvision
                    DoEvents
                    Unload frmBusCoa
                    DoEvents
                    pSetFocus Me.tdbgPDB
                End If
                
    End Select
End Sub

Private Function ValidaMovimiento(cNummov As String, cVoucher As String, cItem As String) As Boolean
    ValidaMovimiento = False
    On Error GoTo serror
    Dim i As Integer
    For i = 0 To lArrDatos.Count(1) - 1
        If CE(lArrDatos(i, 0)) = CE(cNummov) And CE(lArrDatos(i, 1)) = CE(cVoucher) And CE(lArrDatos(i, 7)) = CE(cItem) Then
            Mensajes "Documento ya fue seleccionado en el PDB"
            Exit Function
        End If
    Next i
    
    ValidaMovimiento = True
    Exit Function
serror:
    ValidaMovimiento = False
End Function



Private Sub LlenaProvision()
    On Error GoTo serror
    Dim i As Integer
    Dim Fila As Integer
    Dim nImporte As Integer
    
    i = 0
    Fila = CuentaFilas
    
    If lArrDatos.Count(1) = Fila + 1 Or lArrDatos.Count(1) = Fila Then
        Call AgregaFila
    End If
    
    With frmBusCoa.tdbgProvisiones
        lArrDatos(Fila, 0 + i) = CE(.Columns(0).Value) 'nummov
        lArrDatos(Fila, 1 + i) = CE(.Columns(5).Value) 'voucher
        lArrDatos(Fila, 2 + i) = CE(.Columns(1).Value) 'empresa
        lArrDatos(Fila, 3 + i) = CE(.Columns(2).Value) 'anio
        lArrDatos(Fila, 4 + i) = CE(.Columns(3).Value) 'periodo
        lArrDatos(Fila, 5 + i) = CE(.Columns(4).Value) 'libro
        lArrDatos(Fila, 6 + i) = cTipoPDB  'tipo pdb
        lArrDatos(Fila, 7 + i) = CE(.Columns(21).Value) 'item
        
        lArrDatos(Fila, 8 + i) = "" 'com/ven
        lArrDatos(Fila, 9 + i) = CE(.Columns(11).Value) 'tipo doc
        lArrDatos(Fila, 10 + i) = FE(.Columns(17).Value) 'fecha doc
        
        If CE(.Columns(11).Value) = "12" Then
            lArrDatos(Fila, 11 + i) = ""
        Else
            lArrDatos(Fila, 11 + i) = CE(.Columns(12).Value) 'serie doc
        End If
        
        lArrDatos(Fila, 12 + i) = CE(.Columns(13).Value) 'numero doc
        
        Dim cTipoPersona As String
        cTipoPersona = CE(.Columns(31).Value)  'tipo persona
        
        Select Case cTipoPersona
            Case "N": lArrDatos(Fila, 13 + i) = "01" 'tipo doc ident
            Case "J": lArrDatos(Fila, 13 + i) = "02" 'tipo doc ident
            Case Else: lArrDatos(Fila, 13 + i) = "03" 'tipo doc ident
        End Select
        
        lArrDatos(Fila, 14 + i) = CE(.Columns(32).Value)  'tipo doc ident
        lArrDatos(Fila, 15 + i) = CE(.Columns(9).Value)  'numdoc ident
        lArrDatos(Fila, 16 + i) = "" 'nombre o razon social
        lArrDatos(Fila, 17 + i) = "" 'apellido paterno
        lArrDatos(Fila, 18 + i) = "" 'apellido materno
        lArrDatos(Fila, 19 + i) = "" 'primer nombre
        
        lArrDatos(Fila, 20 + i) = "" 'segundo nombre
        

        lArrDatos(Fila, 21 + i) = "" 'codigo moneda
        lArrDatos(Fila, 22 + i) = "" 'codigo de destino
        lArrDatos(Fila, 23 + i) = "" 'numero de destino
        
        Dim cMoneda As String
        cMoneda = (.Columns(22).Value)   'codigo moneda
        
        Select Case cMoneda
            Case gsMonedaNac: nImporte = NE(.Columns(14).Value) 'soles
            Case Else: nImporte = NE(.Columns(16).Value) 'dolares
        End Select
        
        lArrDatos(Fila, 24 + i) = Redondear(nImporte / 1.19, 2)  ' nBaseImp
        
        If CE(.Columns(11).Value) = "05" Then 'tipo doc
            lArrDatos(Fila, 25 + i) = "" 'nISC
        Else
            lArrDatos(Fila, 25 + i) = 0 'nISC
        End If
        

        lArrDatos(Fila, 26 + i) = Redondear((0.19 / 1.19) * nImporte, 2)   'nIGV
        lArrDatos(Fila, 27 + i) = 0 'nOtros
                       
        lArrDatos(Fila, 28 + i) = "" 'indicador de detracciones
        lArrDatos(Fila, 29 + i) = ""  'codigo de tasa detracciones
        lArrDatos(Fila, 30 + i) = ""  'numero constancia detracciones
        lArrDatos(Fila, 31 + i) = "" 'indicador de retenciones
        lArrDatos(Fila, 32 + i) = "" 'ref tipo doc
        lArrDatos(Fila, 33 + i) = "" 'ref serie doc
        lArrDatos(Fila, 34 + i) = "" 'ref num doc
        lArrDatos(Fila, 35 + i) = "" 'ref fecha doc
        lArrDatos(Fila, 36 + i) = 0 'ref base imp
        lArrDatos(Fila, 37 + i) = 0 'reg igv
        
        lArrDatos(Fila, 38 + i) = CE(.Columns(30).Value) 'medio de pago
        lArrDatos(Fila, 39 + i) = "" 'codigo banco
        lArrDatos(Fila, 40 + i) = "" 'num operacion
        lArrDatos(Fila, 41 + i) = "" 'fecha op
        lArrDatos(Fila, 42 + i) = nImporte 'monto op

    End With
  
    Set tdbgPDB.Array = lArrDatos
    
    Call RebindGrilla(tdbgPDB)
    Call UpdateGrilla(tdbgPDB)

    Call AgregaFila
    
    Exit Sub
serror:
    'Mensajes Err.Description
End Sub

Private Sub AgregaFila()
    Dim Filas As Integer
    Filas = CuentaFilas
    lArrDatos.ReDim 0, Filas, 0, NUM_COLUMNAS    ' filas
    
    Call UpdateGrilla(tdbgPDB)
    Call RebindGrilla(tdbgPDB)
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        tdbgPDB.Width = Me.Width - 200
        tdbgPDB.Height = Me.Height - 1600
    End If
    
    Exit Sub
    
serror:
    
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
 
Private Sub tdbgPDB_BeforeRowColChange(Cancel As Integer)
    With tdbgPDB
'        If .Col = 17 Then
'            If .Columns(.Col) <> "__/__/____" Then
'                ' *** Si fecha no esta completa, completarla
'                '.Columns(.col) = FormatoFecha(.Columns(.col))
'                If VerificaFecha(.Columns(.Col)) = False Then
'                    .RefreshRow
'                    .SetFocus
'                    Cancel = 1
'                    .Columns(.Col) = "__/__/____"
'                End If
'            End If
'        End If
    End With
End Sub

Private Function ValidaPDBReglas() As Boolean
    ValidaPDBReglas = False
    On Error GoTo serror
    Dim i As Integer
    Dim cCadena As String
    Dim cCadena2 As String
    Dim nImporte As Double
    Dim cFecha  As String
    
    Call UpdateGrilla(tdbgPDB)
    
    
    For i = 0 To lArrDatos.Count(1) - 1
        If CE(lArrDatos(i, nCol_cNummov)) <> "" Then
            
            tdbgPDB.Bookmark = i
            
            '---- Regla Campo N.01 ---------'
            If CE(lArrDatos(i, nCol_cTipo)) <> "01" And CE(lArrDatos(i, nCol_cTipo)) <> "02" Then
                Mensajes "TIPO DE COMPRA: El tipo de compra debe ser 01= Compra interna , 02= Compra Externa"
                Exit Function
            End If
            '---- Regla Campo N.02 ---------'
            cCadena = "01,03,04,05,09,07,08,09,10,11,12,13,14,15,16,17,18,21,22,23,24,25,26,27,28,29,30,32,34,35,36,37,87,88"
            
            If CE(lArrDatos(i, nCol_cTipo)) = "01" And _
               InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 0 Then
                Mensajes "TIPO COMP PAGO: Los tipos de documentos validos son: " & cCadena
                Exit Function
            End If
            
            cCadena = "50,52,53,54,91,97,98"
            
            If CE(lArrDatos(i, nCol_cTipo)) = "02" And _
               InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 0 Then
                Mensajes "TIPO COMP PAGO: Los tipos de documentos validos son: " & cCadena
                Exit Function
            End If
            '---- Regla Campo N.03 ---------'
            cCadena = "01,03,04,07,08"
            
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And Len(CE(lArrDatos(i, nCol_cSerieComp))) = 0 Then
                Mensajes "SERIE COMP. PAGO: Ingrese la serie del documento"
                Exit Function
            End If
            
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And Len(CE(lArrDatos(i, nCol_cSerieComp))) > 4 Then
                Mensajes "SERIE COMP. PAGO: Longitud maxima de la serie es 4"
                Exit Function
            End If
            
            
            cCadena = "10,12"
            
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And CE(lArrDatos(i, nCol_cSerieComp)) <> "" Then
                Mensajes "SERIE COMP. PAGO: No se debe ingresar la serie del documento para este tipo de documento"
                Exit Function
            End If
            
            cCadena = "50,52,53,54"
            
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And Len(CE(lArrDatos(i, nCol_cSerieComp))) <> 13 Then
                Mensajes "SERIE COMP. PAGO: Longitud de serie debe ser 13 CCCAAAANNNNNN, CCC=Codigo de Aduana, AAAA=Año, NNNNNN=Num.Correlativo"
                Exit Function
            End If
            
            cCadena = "91,98"
            cCadena2 = "1062,1262,1662"
            
            If CE(lArrDatos(i, nCol_cTipo)) = "02" And _
               InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And _
               InStr(1, cCadena2, CE(lArrDatos(i, nCol_cSerieComp))) = 0 Then
                Mensajes "SERIE COMP. PAGO: Las series de los documentos validos son: " & cCadena2
                Exit Function
            End If
            
            '---- Regla Campo N.04 ---------'
            If Not IsNumeric(lArrDatos(i, nCol_cNumComp)) And CE(lArrDatos(i, nCol_cTipoComp)) <> "12" Then 'diferente de ticket = 12
                Mensajes "NUMERO COMPROBANTE: Numero de documento debe ser numerico"
                Exit Function
            End If
            
            If Val(lArrDatos(i, nCol_cNumComp)) <= 0 Then
                Mensajes "NUMERO COMPROBANTE: Numero de documento debe ser mayor que cero"
                Exit Function
            End If
            
            If Len(CE(lArrDatos(i, nCol_cNumComp))) > 20 Then
                Mensajes "NUMERO COMPROBANTE: Longitud del documento debe ser menor o igual a 20"
                Exit Function
            End If
            
            cCadena = "91,98"
            cCadena2 = "1062,1262,1662"
            
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And InStr(1, cCadena2, CE(lArrDatos(i, nCol_cNumComp))) = 0 Then
                Mensajes "NUMERO COMPROBANTE: Las numeros de los documentos validos son: " & cCadena2
                Exit Function
            End If
            
            cCadena = "50,52,53,54"
                        
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoComp))) = 1 And CE(lArrDatos(i, nCol_cNumComp)) <> "" Then
                Mensajes "NUMERO COMPROBANTE: Las numeros de los documentos deben ser vacios"
                Exit Function
            End If
            
            '---- Regla Campo N.05 ---------'
            cCadena = "01,02,03"
             
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cTipoPer))) = 0 Then
                Mensajes "TIPO PERSONA: Los tipo de persona validos son 01= Natural, 02=Juridico ,03=No Domiciliado"
                Exit Function
            End If
            
            '---- Regla Campo N.06 ---------'
            ' AL JALAR EL DOCUMENTO YA CONVIERTE EL VALOR REQUERIDO
             
            '---- Regla Campo N.07 ---------'
            If Not IsNumeric(lArrDatos(i, nCol_cNumDocPer)) Then
                Mensajes "NUMERO DOC: Numero de documento debe ser numerico"
                Exit Function
            End If
            
            If Val(lArrDatos(i, nCol_cNumDocPer)) <= 0 Then
                Mensajes "NUMERO DOC: Numero de documento debe ser mayor que cero"
                Exit Function
            End If
            
            If Len(CE(lArrDatos(i, nCol_cNumDocPer))) > 12 Then
                Mensajes "NUMERO DOC: Los caracteres del Numero de documento no debe ser mayor a 12"
                Exit Function
            End If
            
            '---- Regla Campo N.08 ---------'
            If CE(lArrDatos(i, nCol_cMedio)) = "" Then
                Mensajes "MEDIO DE PAGO: Ingrese el medio de pago, no debe estar en blanco"
                Exit Function
            End If
            

            nImporte = NE(lArrDatos(i, nCol_nBaseImp)) + NE(lArrDatos(i, nCol_nIGV))
            cCadena = "003,007,009,011"
             
            If nImporte <= 5000 And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 0 Then
                Mensajes "MEDIO DE PAGO: Los medios de pagos permitidos son " & cCadena
                Exit Function
            End If

            '---- Regla Campo N.09 ---------'
            cCadena = "009,011"
             
            If CE(lArrDatos(i, nCol_cCodBan)) <> "" And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 1 Then
                Mensajes "CODIGO DEL BANCO: No se debe ingresar el codigo del banco, para los medios de pago con codigo " & cCadena
                Exit Function
            End If
            
            If CE(lArrDatos(i, nCol_cCodBan)) = "" And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 0 Then
                Mensajes "CODIGO DEL BANCO: Ingrese el codigo del banco"
                Exit Function
            End If

            '---- Regla Campo N.10 ---------'
            cCadena = "009,011"
             
            If CE(lArrDatos(i, nCol_cNumOp)) <> "" And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 1 Then
                Mensajes "NUMERO OPERACION/ NOMBRE BANCO EXT: No se debe ingresar el numero de operacion/ nombre del banco exterior"
                Exit Function
            End If
            
'            If CE(lArrDatos(i, nCol_cNumOp)) = "" Then
'                Mensajes "NUMERO OPERACION/ NOMBRE BANCO EXT: Ingrese el numero de operacion/ nombre del banco exterior"
'                Exit Function
'            End If
            
            '---- Regla Campo N.11 ---------'
            cCadena = "009"
             
            If CE(lArrDatos(i, nCol_cFecOp)) <> "" And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 1 Then
                Mensajes "FECHA OPERACION: No se debe ingresar la fecha de operacion"
                Exit Function
            End If
            
             If IsDate(lArrDatos(i, nCol_cFecOp)) = False And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 0 Then
                Mensajes "FECHA OPERACION: Fecha de operacion no valida"
                Exit Function
            End If
            
             If CE(lArrDatos(i, nCol_cFecOp)) = "" And InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 0 Then
                Mensajes "FECHA OPERACION: Ingrese la fecha de operacion"
                Exit Function
            End If
            
            
            cFecha = FE(lArrDatos(i, nCol_dFechaComp))
            If InStr(1, cCadena, CE(lArrDatos(i, nCol_cMedio))) = 0 Then
                If CDate(FE(lArrDatos(i, nCol_cFecOp))) < CDate(cFecha) Then
                    Mensajes "FECHA OPERACION: Fecha de operacion no debe ser menor a la fecha de emision del comprobante de pago"
                    Exit Function
                End If
                
                If Month(CDate(lArrDatos(i, nCol_cFecOp))) <> Val(tdbcMes.BoundText) Then
                    Mensajes "FECHA OPERACION: La fecha de operacion debe estar en el mes seleccionado"
                    Exit Function
                End If
                
                If Year(CDate(lArrDatos(i, nCol_cFecOp))) <> Val(gsAnio) Then
                    Mensajes "FECHA OPERACION: La fecha de operacion debe estar en el año del sistema"
                    Exit Function
                End If
            End If
            '---- Regla Campo N.12 ---------'
            If NE(lArrDatos(i, nCol_nMonOp)) > nImporte Then
                Mensajes "MONTO OPERACION: El monto de operacion no debe pasar el total de adquisicion " & Format(CE(nImporte), "###,###,##0.00")
                Exit Function
            End If
            
        End If
    Next i
    
    ValidaPDBReglas = True

    Exit Function
serror:
    ValidaPDBReglas = False
End Function


Private Sub cmdPreliminar_Click()
    Dim matriz(8) As Variant
    Dim Titulo As String
    cmdPreliminar.Enabled = False
    DoEvents
    Titulo = "Reporte de PDB - Pagos"
    Titulo = UCase(Titulo)
    matriz(0) = "@Accion;BUSCARTODOS;True"
    matriz(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(2) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
    matriz(4) = "@Lib_cTipoLibro;" & tdbcLibro.BoundText & ";True"
    matriz(5) = "@Dco_cTipoPDB;P;True"
    
    matriz(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(7) = "@RUC;" & "RUC : " & gsRUC & ";True"
    matriz(8) = "@NOMBREMES;" & NombreMes(tdbcMes.BoundText) & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptPDBPagos.rpt", crptToWindow, Titulo, "", matriz(), formulas()
  
    cmdPreliminar.Enabled = True
End Sub
