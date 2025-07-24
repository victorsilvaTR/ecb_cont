VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManFlujosCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros de cuentas adicionales para Flujo Efectivo"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   Icon            =   "frmManFlujosCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   7095
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7035
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1215
         TabIndex        =   13
         Top             =   270
         Width           =   3150
         _ExtentX        =   5556
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
         _PropDict       =   $"frmManFlujosCuentas.frx":0ECA
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   0
      TabIndex        =   7
      Top             =   630
      Width           =   7035
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1215
         MaxLength       =   12
         TabIndex        =   9
         Top             =   225
         Width           =   3090
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   735
      End
      Begin MSForms.CommandButton cmdActualiza 
         Height          =   390
         Left            =   4410
         TabIndex        =   8
         Top             =   180
         Width           =   1290
         Caption         =   " Buscar"
         PicturePosition =   327683
         Size            =   "2275;688"
         Picture         =   "frmManFlujosCuentas.frx":0F51
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin TrueOleDBGrid70.TDBGrid grdEgresos 
      Height          =   3840
      Left            =   45
      TabIndex        =   0
      Top             =   1845
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cuenta"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TITULO"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "SUBTITULO"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4551"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4471"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(28)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(29)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
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
      DeadAreaBackColor=   12632256
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
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=43,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=25,.alignment=0"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=26"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=27"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=43"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=26"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=27"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=16,.parent=25"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=26"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=27"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=43"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=20,.parent=25"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=26"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=27"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=43"
      _StyleDefs(53)  =   "Named:id=33:Normal"
      _StyleDefs(54)  =   ":id=33,.parent=0"
      _StyleDefs(55)  =   "Named:id=34:Heading"
      _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   ":id=34,.wraptext=-1"
      _StyleDefs(58)  =   "Named:id=35:Footing"
      _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   "Named:id=36:Selected"
      _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(62)  =   "Named:id=37:Caption"
      _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(64)  =   "Named:id=38:HighlightRow"
      _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(66)  =   "Named:id=39:EvenRow"
      _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(68)  =   "Named:id=40:OddRow"
      _StyleDefs(69)  =   ":id=40,.parent=33"
      _StyleDefs(70)  =   "Named:id=41:RecordSelector"
      _StyleDefs(71)  =   ":id=41,.parent=34"
      _StyleDefs(72)  =   "Named:id=42:FilterBar"
      _StyleDefs(73)  =   ":id=42,.parent=33"
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   1620
      TabIndex        =   6
      ToolTipText     =   "Grabar modificaciones"
      Top             =   1395
      Width           =   1380
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManFlujosCuentas.frx":14EB
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdListar 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Cargar nueva Configuración"
      Top             =   1395
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujosCuentas.frx":1A85
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   390
      Left            =   5760
      TabIndex        =   4
      Top             =   1395
      Width           =   1290
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2275;688"
      Picture         =   "frmManFlujosCuentas.frx":201F
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOrdenar 
      Height          =   390
      Left            =   4410
      TabIndex        =   3
      Top             =   1395
      Width           =   1290
      Caption         =   " Ordenar Item"
      PicturePosition =   327683
      Size            =   "2275;688"
      Picture         =   "frmManFlujosCuentas.frx":25B9
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   390
      Left            =   3060
      TabIndex        =   2
      Top             =   1395
      Width           =   1290
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2275;688"
      Picture         =   "frmManFlujosCuentas.frx":2B53
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   45
      Index           =   0
      Left            =   8010
      TabIndex        =   1
      Top             =   4230
      Width           =   300
   End
End
Attribute VB_Name = "frmManFlujosCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrPspto As New XArrayDB
Dim valorPres As Double
Dim Fila As Integer
Dim lTipoPres As String
Dim lArrDet() As Variant
Dim lControl As String
Dim nFilas As Integer
Dim TCMensual(12) As Double
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdActualiza_Click()
    Call LlamaBuscar(frmBuscador, "Cuenta", "Cuenta", "CuentasNo2D", Me, gsPeriodo, txtCuenta.Text)
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Dim i As Integer
    Dim psql As String
    Dim Valor As String
    Dim Titulo As String
    Dim STitulo As String
    
    Titulo = Left(param2, 1)
    STitulo = Right(param2, 1)
    
    Valor = True
    Select Case lControl
            Case "CuentasNo2D"
                ' *** Ver si centro de costo no esta en el Grid
                For i = 0 To lArrPspto.Count(1) - 1
                    If Trim(lArrPspto(i, 0)) = Trim(param0) Then
                        Valor = False
                        Exit For
                    End If
                Next
                ' ***
                Dim Fila As Integer
                Dim Filas As Integer
                
                If Valor = True Then
                    
                    Fila = lArrPspto.Count(1) - 1
                    Filas = lArrPspto.Count(1)
                    
                    lArrPspto.ReDim 0, Filas, 0, 3
                    
                    If Fila < 0 Then Fila = 0
                    'On Error Resume Next
                    If CE(lArrPspto(Fila, 0)) <> "" Then
                        Fila = Fila + 1
                    End If
                    
                    If Fila > Filas Then
                        Filas = Filas + 1
                    End If
                    
                    lArrPspto.ReDim 0, Filas, 0, 13
                    lArrPspto(Fila, 0) = CE(param0)
                    lArrPspto(Fila, 1) = CE(param1)
                    lArrPspto(Fila, 2) = "N"
                    lArrPspto(Fila, 3) = "N"

                    
                    Set grdEgresos.Array = lArrPspto
                    grdEgresos.ReBind
                    Unload frmBuscador
                    
                    'txtCuenta.Text = ""
                    'pSetFocus txtCuenta
                    
                Else
                    Mensajes "La cuenta seleccionada, ya esta contenido actualmente", vbInformation
                End If
    End Select
End Sub

Private Sub cmdEliminaItem_Click()
    Dim i As Integer
    Dim cad1 As String
    Dim cad2 As String
    Dim Valor As Boolean
    
    Valor = True
    If lArrPspto.Count(1) = 1 And grdEgresos.Bookmark = 0 And CE(grdEgresos.Columns(0)) = "" Then
        Exit Sub
    End If
    
    
    If lArrPspto.Count(1) > 0 Then
        If Valor = True And Not IsNull(grdEgresos.Bookmark) Then
            lArrPspto.DeleteRows (grdEgresos.Bookmark)
            grdEgresos.ReBind
            pSetFocus grdEgresos
        End If
        
        If lArrPspto.Count(1) = 0 Then
            lArrPspto.Clear
        End If
        

    End If
End Sub

Private Function Grabar() As Boolean
    On Error GoTo serror
    Grabar = False

    If (lArrPspto.Count(1) = 1 Or lArrPspto.Count(2) = 1) And grdEgresos.Bookmark = 0 And CE(grdEgresos.Columns(0)) = "" Then
        Exit Function
    End If

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    
    Set clsMante = New clsMantoTablas
    
    ReDim lArrDet(3) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcMes.BoundText
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "SpCn_FlujoCuentas", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Function
    End If
    
    For i = 0 To lArrPspto.Count(1) - 1
        If CE(lArrPspto(i, 0)) <> "" Then
            Call CargaArregloDet(i)
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "SpCn_FlujoCuentas", lArrDet(), False) = False Then
                Screen.MousePointer = vbNormal
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                Exit Function
            End If
        End If
    Next
    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal
    
    Set clsMante = Nothing
    
    Me.cmdEliminaItem.Enabled = True
    Me.cmdActualiza.Enabled = True
    DoEvents

    Call GeneraArreglo
    
    
    Grabar = True
    Exit Function
serror:
    Mensajes Err.Description
    Grabar = False
End Function

Private Sub CargaArregloDet(item As Integer)
    ReDim lArrDet(4) As Variant
    lArrDet(0) = "INSERTAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcMes.BoundText
    lArrDet(4) = lArrPspto(item, 0)
End Sub


Private Sub cmdGrabar_Click()
    If Grabar = True Then
       On Error Resume Next
       grdEgresos.Update
       DoEvents
       Mensajes "Se grabaron las cuentas ", vbInformation
       
       grdEgresos.Refresh
       
    End If

End Sub

Private Sub cmdListar_Click()
    Call GeneraArreglo
End Sub

Private Sub cmdOrdenar_Click()
    Grabar
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call LlenaComboMesApeAddItem(tdbcMes)
    
    tdbcMes.BoundText = gsPeriodo
    
    DoEvents
    
    tdbcMes.ReBind
    
    Dim sqlcombos As String
     grdEgresos.FetchRowStyle = True
    
    Me.Top = (frmMDIConta.ScaleHeight - Me.Height) / 2
    Me.Left = (frmMDIConta.ScaleWidth - Me.Width) / 2
    

    Dim strMon As String
    If gsByMoneda = 1 Then
        strMon = " (Mon_cMNac = '1' or Mon_cMExt = '1') "
    Else
        strMon = " Mon_cMNac = '1' "
    End If
    
    Call GeneraArreglo
    grdEgresos.Splits(0).MarqueeStyle = dbgHighlightRow
    grdEgresos.HighlightRowStyle = "HighlightRow"
    
    grdEgresos.ReBind

    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdEliminaItem.Enabled = False
        cmdActualiza.Enabled = False
        cmdOrdenar.Enabled = False
        cmdGrabar.Enabled = False
    Else
        cmdEliminaItem.Enabled = True
        cmdActualiza.Enabled = True
        cmdOrdenar.Enabled = True
        cmdGrabar.Enabled = True
        
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
        Case 115: If cmdGrabar.Enabled Then Grabar
        Case 118:
    End Select
    ' ***
End Sub

Private Sub GeneraArreglo()
    ' ***
    Dim sqlPres As String
    Dim i As Integer
    
    lTipoPres = "I"
    On Local Error GoTo ErrorEjecucion
    sqlPres = "SpCn_FlujoCuentas 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "'"
    
    Call GridArreglo(lArrPspto, Me.grdEgresos, sqlPres)
   
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub



Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
       
        With grdEgresos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Me.Width - .Left - 200
            .Height = Me.Height - .Top - 500
        End With
        
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lArrPspto = Nothing
End Sub

Private Sub grdEgresos_AfterColEdit(ByVal ColIndex As Integer)
    Dim Fila As Integer
    If ColIndex >= 3 Then
        If grdEgresos.Columns(ColIndex) = "" Then grdEgresos.Columns(ColIndex) = 0
        lArrPspto(grdEgresos.Bookmark, 3) = (grdEgresos.Columns(3) - valorPres) + grdEgresos.Columns(ColIndex)
        grdEgresos.Columns(3) = lArrPspto(grdEgresos.Bookmark, 3)
        grdEgresos.Update
        
        
    End If

End Sub

Private Sub grdEgresos_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If CE(lArrPspto(grdEgresos.Bookmark, 2)) = "S" Or CE(lArrPspto(grdEgresos.Bookmark, 18)) = "S" Then
        Cancel = 1
        Exit Sub
    
    End If
    
    If ColIndex = 3 Then
        Cancel = 1
        Exit Sub
    End If
    
End Sub

Private Sub grdEgresos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    
    If lArrPspto.Count(1) > 1 Then
        If CE(lArrPspto(Bookmark, 2)) = "S" Then 'titulo
            RowStyle.BackColor = gsColorCCTitulo
        End If
        If CE(lArrPspto(Bookmark, 3)) = "S" Then 'subtitulo
            RowStyle.BackColor = gsColorCCSTitulo
        End If
        
    End If
End Sub

Private Sub optEgresos_Click()
    Call GeneraArreglo
End Sub

Private Sub optIngresos_Click()
    Call GeneraArreglo
End Sub

Private Sub tdbcMes_ItemChange()
    cmdListar_Click
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdActualiza_Click
    End If
End Sub
