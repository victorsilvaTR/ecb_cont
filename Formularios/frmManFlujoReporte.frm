VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManFlujoReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del Reporte del Flujo de Efectivo "
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "frmManFlujoReporte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   11460
   Begin TrueOleDBGrid70.TDBDropDown tdbdTipo 
      Height          =   1605
      Left            =   7335
      TabIndex        =   14
      Top             =   3285
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   2831
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
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
   Begin VB.Frame fraImportar 
      Caption         =   " Importar datos del periodo seleccionado "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   5535
      TabIndex        =   7
      Top             =   45
      Width           =   5775
      Begin TrueOleDBList70.TDBCombo tdbcMesImportar 
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   315
         Width           =   3930
         _ExtentX        =   6932
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
         _PropDict       =   $"frmManFlujoReporte.frx":0ECA
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
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Width           =   660
      End
      Begin MSForms.CommandButton cmdImportar 
         Height          =   360
         Left            =   4050
         TabIndex        =   9
         ToolTipText     =   " Importar los datos del periodo seleccionado "
         Top             =   720
         Width           =   1305
         Caption         =   " Importar"
         PicturePosition =   327683
         Size            =   "2302;635"
         Picture         =   "frmManFlujoReporte.frx":0F51
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo de configuracion del reporte "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   5010
      Begin VB.ComboBox cboMetodo 
         Height          =   315
         ItemData        =   "frmManFlujoReporte.frx":14EB
         Left            =   1035
         List            =   "frmManFlujoReporte.frx":14F5
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   3750
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1035
         TabIndex        =   5
         Top             =   315
         Width           =   3750
         _ExtentX        =   6615
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
         _PropDict       =   $"frmManFlujoReporte.frx":150D
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
         Caption         =   "Metodo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   765
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   705
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgFlujo 
      Height          =   4110
      Left            =   90
      TabIndex        =   13
      Top             =   1845
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   7250
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Item"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "tdbdPlantilla"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "CodTipo"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo de Formula"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tdbdTipo"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "FormulaBD"
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "tdbdColumna"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Formula o  Valor"
      Columns(5).DataField=   ""
      Columns(5).DropDown=   "tdbdColumna"
      Columns(5).DropDown.vt=   8
      Columns(5).ButtonPicture.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ButtonPicture(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
      Columns(5).ButtonPicture(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(2)=   "7+vx7+vx7+vx7+vx7+trrYQhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(3)=   "7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU"
      Columns(5).ButtonPicture(4)=   "3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx"
      Columns(5).ButtonPicture(5)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(6)=   "7+vx7+vx7+vx7+trrYQhhCkhhCkhhCkhhCkhhCmU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx"
      Columns(5).ButtonPicture(7)=   "7+tjpWM5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVIhhCnx7+vx7+tjpWOU3oyU"
      Columns(5).ButtonPicture(8)=   "3oyU3oyU3oyU3oyU3ow5tVKU3oyU3oyU3oyU3oyU3owhhCnx7+vx7+trrYRjpWNjpWNjpWNjpWNj"
      Columns(5).ButtonPicture(9)=   "pWOU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIh"
      Columns(5).ButtonPicture(10)=   "hCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx"
      Columns(5).ButtonPicture(11)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(12)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(13)=   "7+vx7+vx7+vx7+trrYRjpWNjpWNrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(5).ButtonPicture(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+s="
      Columns(5).ButtonPicture.vt=   9
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=139777"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=8229"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=8149"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(1).DropDownList=1"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=139780"
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(29)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(30)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(32)=   "Column(3).AutoDropDown=1"
      Splits(0)._ColumnProps(33)=   "Column(3).DropDownList=1"
      Splits(0)._ColumnProps(34)=   "Column(4).Width=1244"
      Splits(0)._ColumnProps(35)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._WidthInPix=1164"
      Splits(0)._ColumnProps(37)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(38)=   "Column(4)._ColStyle=131588"
      Splits(0)._ColumnProps(39)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(41)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(42)=   "Column(5).Width=7646"
      Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=7567"
      Splits(0)._ColumnProps(45)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=131588"
      Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(50)=   "Column(5).AutoDropDown=1"
      Splits(0)._ColumnProps(51)=   "Column(5).DropDownList=1"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   688
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   14215660
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=6"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1111"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=139777"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=8229"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=8149"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=139780"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(1).DropDownList=1"
      Splits(1)._ColumnProps(18)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(19)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(20)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(21)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(22)=   "Column(2)._ColStyle=139780"
      Splits(1)._ColumnProps(23)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(24)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(25)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(26)=   "Column(3).Width=2196"
      Splits(1)._ColumnProps(27)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(28)=   "Column(3)._WidthInPix=2117"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=131588"
      Splits(1)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(31)=   "Column(3).AutoDropDown=1"
      Splits(1)._ColumnProps(32)=   "Column(3).DropDownList=1"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=1244"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=1164"
      Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=131588"
      Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(41)=   "Column(5).Width=529"
      Splits(1)._ColumnProps(42)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(5)._WidthInPix=450"
      Splits(1)._ColumnProps(44)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(45)=   "Column(5)._ColStyle=131588"
      Splits(1)._ColumnProps(46)=   "Column(5).Button=1"
      Splits(1)._ColumnProps(47)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(48)=   "Column(5).AutoDropDown=1"
      Splits(1)._ColumnProps(49)=   "Column(5).DropDownList=1"
      Splits(1)._ColumnProps(50)=   "Column(5).ButtonAlways=1"
      Splits.Count    =   2
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
      _StyleDefs(25)  =   "Splits(0).Style:id=21,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=44,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=22,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=23,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=24,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=26,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=25,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=27,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=28,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=43,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=45,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=46,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=62,.parent=21,.alignment=2,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=22"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=23"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=25"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=66,.parent=21,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=22"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=23"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=25"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=106,.parent=21,.locked=-1"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=22"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=23"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=25"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=74,.parent=21,.bgcolor=&HFFFFFF&"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=22"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=23"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=25"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=78,.parent=21,.bgcolor=&HFFFFFF&"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=22"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=23"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=25"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=94,.parent=21,.bgcolor=&HFFFFFF&"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=91,.parent=22"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=92,.parent=23"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=93,.parent=25"
      _StyleDefs(61)  =   "Splits(1).Style:id=79,.parent=1"
      _StyleDefs(62)  =   "Splits(1).CaptionStyle:id=88,.parent=4"
      _StyleDefs(63)  =   "Splits(1).HeadingStyle:id=80,.parent=2"
      _StyleDefs(64)  =   "Splits(1).FooterStyle:id=81,.parent=3"
      _StyleDefs(65)  =   "Splits(1).InactiveStyle:id=82,.parent=5"
      _StyleDefs(66)  =   "Splits(1).SelectedStyle:id=84,.parent=6"
      _StyleDefs(67)  =   "Splits(1).EditorStyle:id=83,.parent=7"
      _StyleDefs(68)  =   "Splits(1).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(69)  =   "Splits(1).EvenRowStyle:id=86,.parent=9"
      _StyleDefs(70)  =   "Splits(1).OddRowStyle:id=87,.parent=10"
      _StyleDefs(71)  =   "Splits(1).RecordSelectorStyle:id=89,.parent=11"
      _StyleDefs(72)  =   "Splits(1).FilterBarStyle:id=90,.parent=12"
      _StyleDefs(73)  =   "Splits(1).Columns(0).Style:id=32,.parent=79,.alignment=2,.locked=-1"
      _StyleDefs(74)  =   "Splits(1).Columns(0).HeadingStyle:id=29,.parent=80"
      _StyleDefs(75)  =   "Splits(1).Columns(0).FooterStyle:id=30,.parent=81"
      _StyleDefs(76)  =   "Splits(1).Columns(0).EditorStyle:id=31,.parent=83"
      _StyleDefs(77)  =   "Splits(1).Columns(1).Style:id=50,.parent=79,.locked=-1"
      _StyleDefs(78)  =   "Splits(1).Columns(1).HeadingStyle:id=47,.parent=80"
      _StyleDefs(79)  =   "Splits(1).Columns(1).FooterStyle:id=48,.parent=81"
      _StyleDefs(80)  =   "Splits(1).Columns(1).EditorStyle:id=49,.parent=83"
      _StyleDefs(81)  =   "Splits(1).Columns(2).Style:id=110,.parent=79,.locked=-1"
      _StyleDefs(82)  =   "Splits(1).Columns(2).HeadingStyle:id=107,.parent=80"
      _StyleDefs(83)  =   "Splits(1).Columns(2).FooterStyle:id=108,.parent=81"
      _StyleDefs(84)  =   "Splits(1).Columns(2).EditorStyle:id=109,.parent=83"
      _StyleDefs(85)  =   "Splits(1).Columns(3).Style:id=20,.parent=79,.bgcolor=&HFFFFFF&"
      _StyleDefs(86)  =   "Splits(1).Columns(3).HeadingStyle:id=17,.parent=80"
      _StyleDefs(87)  =   "Splits(1).Columns(3).FooterStyle:id=18,.parent=81"
      _StyleDefs(88)  =   "Splits(1).Columns(3).EditorStyle:id=19,.parent=83"
      _StyleDefs(89)  =   "Splits(1).Columns(4).Style:id=54,.parent=79,.bgcolor=&HFFFFFF&"
      _StyleDefs(90)  =   "Splits(1).Columns(4).HeadingStyle:id=51,.parent=80"
      _StyleDefs(91)  =   "Splits(1).Columns(4).FooterStyle:id=52,.parent=81"
      _StyleDefs(92)  =   "Splits(1).Columns(4).EditorStyle:id=53,.parent=83"
      _StyleDefs(93)  =   "Splits(1).Columns(5).Style:id=58,.parent=79,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(94)  =   "Splits(1).Columns(5).HeadingStyle:id=55,.parent=80"
      _StyleDefs(95)  =   "Splits(1).Columns(5).FooterStyle:id=56,.parent=81"
      _StyleDefs(96)  =   "Splits(1).Columns(5).EditorStyle:id=57,.parent=83"
      _StyleDefs(97)  =   "Named:id=33:Normal"
      _StyleDefs(98)  =   ":id=33,.parent=0"
      _StyleDefs(99)  =   "Named:id=34:Heading"
      _StyleDefs(100) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   ":id=34,.wraptext=-1"
      _StyleDefs(102) =   "Named:id=35:Footing"
      _StyleDefs(103) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   "Named:id=36:Selected"
      _StyleDefs(105) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(106) =   "Named:id=37:Caption"
      _StyleDefs(107) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(108) =   "Named:id=38:HighlightRow"
      _StyleDefs(109) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(110) =   "Named:id=39:EvenRow"
      _StyleDefs(111) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(112) =   "Named:id=40:OddRow"
      _StyleDefs(113) =   ":id=40,.parent=33"
      _StyleDefs(114) =   "Named:id=41:RecordSelector"
      _StyleDefs(115) =   ":id=41,.parent=34"
      _StyleDefs(116) =   "Named:id=42:FilterBar"
      _StyleDefs(117) =   ":id=42,.parent=33"
   End
   Begin MSForms.ToggleButton cmdVisibleImport 
      Height          =   375
      Left            =   5130
      TabIndex        =   12
      ToolTipText     =   " Importar configuración "
      Top             =   135
      Width           =   375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "661;661"
      Value           =   "0"
      Picture         =   "frmManFlujoReporte.frx":1594
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   8325
      TabIndex        =   11
      Top             =   1395
      Width           =   1575
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoReporte.frx":1B2E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   4905
      TabIndex        =   3
      ToolTipText     =   "Eliminar la cuenta selecccionada "
      Top             =   1395
      Width           =   1575
      Caption         =   " Eliminar Cuentas"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoReporte.frx":20C8
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   6615
      TabIndex        =   2
      ToolTipText     =   "Eliminar todas las cuentas de la lista "
      Top             =   1395
      Width           =   1575
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoReporte.frx":2662
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   3195
      TabIndex        =   1
      ToolTipText     =   "Graba la lista mostrada "
      Top             =   1395
      Width           =   1575
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoReporte.frx":2BFC
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   375
      Left            =   1485
      TabIndex        =   0
      ToolTipText     =   " Vuelve a cargar los datos almacenados "
      Top             =   1395
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoReporte.frx":3196
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmManFlujoReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lArrFlujo As New XArrayDB
Dim lArrDetalle(8) As Variant
Dim rsTipo As ADODB.Recordset

Dim gsGrupo As String
Dim gsColumna As Integer
Const NUM_COL = 6

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Sub cboMetodo_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdEliminaItem_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If

    cmdEliminaItem.Enabled = False
    DoEvents
    
    If CE(tdbgFlujo.Columns(0).Value) = "" Then
           On Error Resume Next
           lArrFlujo.DeleteRows tdbgFlujo.Bookmark
           UpdateGrilla
           tdbgFlujo.ReBind
    Else
        If MsgBox("Deseas eliminar la fila seleccionada", vbYesNo + vbQuestion) = vbYes Then
            
           On Error Resume Next
           tdbgFlujo.Columns(2) = ""
           tdbgFlujo.Columns(3) = ""
           tdbgFlujo.Columns(4) = ""
           tdbgFlujo.Columns(5) = ""
           'tdbgFlujo.Columns(6) = ""
           'tdbgFlujo.Columns(7) = ""

           UpdateGrilla
        End If
    End If
    
    cmdEliminaItem.Enabled = True
End Sub

Private Sub cmdEliminarTodo_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If

    If MsgBox("Deseas eliminar todas las cuentas de la lista", vbYesNo + vbQuestion) = vbYes Then
    Dim i As Integer
        For i = 0 To lArrFlujo.Count(1) - 1
            lArrFlujo(i, 2) = ""
            lArrFlujo(i, 3) = ""
            lArrFlujo(i, 4) = ""
            lArrFlujo(i, 5) = ""
        Next i
    
       tdbgFlujo.Refresh
       
       UpdateGrilla
    End If

End Sub


Private Function CargaArregloDet(item As Integer) As Boolean
    CargaArregloDet = True
    
    lArrDetalle(0) = "INSERTAR"
    lArrDetalle(1) = gsEmpresa
    lArrDetalle(2) = gsAnio
    lArrDetalle(3) = tdbcMes.BoundText
    
    lArrDetalle(4) = CE(lArrFlujo(item, 0))
    lArrDetalle(5) = CE(lArrFlujo(item, 2))
    lArrDetalle(6) = CE(lArrFlujo(item, 4))
    
    If CE(lArrFlujo(item, 2)) = "M" Then
       lArrDetalle(6) = CE(lArrFlujo(item, 5))
    End If
    
    lArrDetalle(7) = CE(lArrFlujo(item, 5))
    lArrDetalle(8) = CE(cboMetodo.Text)  'METODO
End Function

Private Sub Grabar()
 
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If
    
    If ValidaCampos = False Then
        Exit Sub
    End If
    
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas
    

    Dim lArrDet(8) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcMes.BoundText
    
    lArrDet(4) = ""
    lArrDet(5) = ""
    lArrDet(6) = ""
    lArrDet(7) = ""
    
    lArrDet(8) = cboMetodo.Text
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoReporte", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To lArrFlujo.Count(1) - 1
        If CE(lArrFlujo(i, 2)) <> "" Then
            
                If CargaArregloDet(i) = True Then
                    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoReporte", lArrDetalle(), False) = False Then
                        Screen.MousePointer = vbNormal
                        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                        Exit Sub
                    End If
                End If
            
        End If
    Next

    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal

    Set clsMante = Nothing
    
    DoEvents
    
    cmdRefresh_Click
    
    DoEvents
    Mensajes "Se ha grabado con exito ...", vbInformation
   
    
End Sub

Private Sub cmdGrabar_Click()
    Grabar
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
       Exit Function
    End If
    
    If (lArrFlujo.Count(1) = 1 Or lArrFlujo.Count(2) = 1) And tdbgFlujo.Bookmark = 0 And CE(tdbgFlujo.Columns(0)) = "" Then
       ValidaCampos = True
       Exit Function
    End If
    
    Dim i As Integer
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If CE(lArrFlujo(i, 3)) <> "" And CE(lArrFlujo(i, 5)) = "" Then
           Mensajes "Ingrese una formula o valor, para la cuenta " & lArrFlujo(i, 0)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = 5
           
           pSetFocus tdbgFlujo
           Exit Function
        End If
    Next i
    
    ValidaCampos = True
End Function

Private Sub cmdImportar_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
        pSetFocus tdbcMes
       Exit Sub
    End If
    
    If tdbcMesImportar.Text = "" Then
        Mensajes "Seleccione el periodo de importación"
        pSetFocus tdbcMesImportar
       Exit Sub
    End If
    
    If cboMetodo.Text = "" Then
        Mensajes "Seleccione el metodo"
        pSetFocus cboMetodo
       Exit Sub
    End If
    
    
    If tdbcMes.BoundText = tdbcMesImportar.BoundText Then
        Mensajes "El periodo del proceso debe ser diferente al periodo de importación"
        pSetFocus tdbcMes
       Exit Sub
    End If

    If Mensajes("Desea importar los datos del mes seleccionado", vbYesNo + vbQuestion) = vbYes Then
        cmdImportar.Enabled = False
        Screen.MousePointer = vbHourglass
        
        GeneraArreglo tdbcMesImportar.BoundText
        DoEvents
        cmdImportar.Enabled = True
        Screen.MousePointer = vbNormal
        
        cmdVisibleImport.Value = False
        
        tdbgFlujo.ReBind
        On Error Resume Next
        tdbgFlujo.Row = 0
        tdbgFlujo.Bookmark = 0
    End If
End Sub


Private Sub LlenaCombos()
    Dim sql As String
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    DoEvents
    tdbcMes.BoundText = gsPeriodo
    
    Call LlenaComboMesApeAddItem(tdbcMesImportar)
    DoEvents
    tdbcMesImportar.BoundText = gsPeriodo
   
End Sub

Private Sub LlenaListas()
    Dim Codigo As String
    Dim sql As String
    '--------------------------------------------------
    sql = "SELECT TAB_CDESCRIPCAMPO , TAB_CCODIGO " & _
          "FROM TABLA WHERE TAB_CTABLA='070' AND EMP_CCODIGO='" & gsEmpresa & "' and TAB_CCODIGO<>'C' " & _
          "ORDER BY TAB_CDESCRIPCAMPO "
    
    Call CerrarRecordSet(rsTipo)
    Call LlenarRecordSet(sql, rsTipo)
    
    Set tdbdTipo.DataSource = rsTipo
    
End Sub

Private Sub GeneraArreglo(Mes As String)
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion
    
    If tdbcMes.Text <> "" And cboMetodo.Text = "" Then
        Mensajes "Seleccione el metodo"
        pSetFocus cboMetodo
        Exit Sub
    End If
    
    
    If tdbcMes.Text <> "" And cboMetodo.Text <> "" Then
        Set lArrFlujo = New XArrayDB
        sql = "spCn_FlujoReporte 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & Mes & "','','','','','" & cboMetodo.Text & "'"
        Call GridArreglo(lArrFlujo, tdbgFlujo, sql)
        
        If lArrFlujo.Count(2) < NUM_COL Then
           lArrFlujo.ReDim 0, Filas, 0, NUM_COL
        End If
    Else
        lArrFlujo.ReDim 0, 0, 1, NUM_COL
    End If
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub

Private Function CuentaFilas() As Integer
    Dim i As Integer
    Dim Contador As Integer
    Contador = 0
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If CE(lArrFlujo(i, 0)) <> "" Then
           Contador = Contador + 1
        End If
    Next i
    
    CuentaFilas = Contador
End Function


Private Sub cmdRefresh_Click()
    On Error GoTo serror:
    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    
    LlenaListas
    
    DoEvents
    
    GeneraArreglo tdbcMes.BoundText
    
    DoEvents
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
serror:
    
    On Error Resume Next
    tdbgFlujo.ReBind
    tdbgFlujo.Row = 0
    tdbgFlujo.Bookmark = 0
    
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdVisibleImport_Click()
    If cmdVisibleImport.Value = True Then
        fraImportar.Visible = True
    Else
        fraImportar.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
                If MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar") = vbYes Then
                    Unload Me
                End If

        Case 115: If cmdGrabar.Enabled Then cmdGrabar_Click

    End Select
End Sub

Private Sub Form_Load()
    Me.Height = 6525
    Me.Width = 11550
    
    tdbgFlujo.FetchRowStyle = True
    tdbgFlujo.Splits(0).MarqueeStyle = dbgHighlightRow
    tdbgFlujo.Splits(1).MarqueeStyle = dbgFloatingEditor
    LlenaCombos
    DoEvents
    cmdRefresh_Click
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabar.Enabled = False
        cmdEliminaItem.Enabled = False
        cmdEliminarTodo.Enabled = False
        cmdImportar.Enabled = False
        tdbgFlujo.Splits(0).Locked = True
        tdbgFlujo.Splits(1).Locked = True
        tdbgFlujo.Splits(1).Columns(5).Button = False

        cmdVisibleImport.Enabled = False
    Else
        cmdGrabar.Enabled = True
        cmdEliminaItem.Enabled = True
        cmdEliminarTodo.Enabled = True
        cmdImportar.Enabled = True
        tdbgFlujo.Splits(0).Locked = False
        tdbgFlujo.Splits(1).Locked = False
        tdbgFlujo.Splits(1).Columns(5).Button = True
        
        cmdVisibleImport.Enabled = True
    End If
    
    cmdVisibleImport_Click
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        tdbgFlujo.Height = Me.Height - 2000 - 300
        tdbgFlujo.Width = Me.Width - 300
    End If
    
    If Me.WindowState = vbMaximized Then
        tdbgFlujo.Columns(1).Width = 4890
    End If
    
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        tdbgFlujo.Columns(1).Width = 3540
    End If
    
    Exit Sub
    
serror:

End Sub

Private Sub tdbcMes_ItemChange()
    If tdbcMes.Text = "" Then
        tdbgFlujo.Splits(0).Locked = True
    Else
        tdbgFlujo.Splits(0).Locked = False
    End If
    
    cmdRefresh_Click
End Sub

Private Sub tdbdTipo_DropDownClose()
    tdbgFlujo.Columns(2) = tdbdTipo.Columns(1).Value
    tdbgFlujo.Columns(3) = tdbdTipo.Columns(0).Value
    tdbgFlujo.Columns(4) = ""
    tdbgFlujo.Columns(5) = ""
    DoEvents
    On Error Resume Next
    tdbgFlujo.Update
    DoEvents
    pSetFocus tdbgFlujo

End Sub

Private Sub tdbgFlujo_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If lArrFlujo Is Nothing Or IsNull(tdbgFlujo.Bookmark) Then
        Exit Sub
    End If
    If lArrFlujo.Count(2) >= NUM_COL Then
        If CE(lArrFlujo(tdbgFlujo.Bookmark, 6)) = "S" Or CE(lArrFlujo(tdbgFlujo.Bookmark, 6)) = "T" Then
           Cancel = 1
        End If
    End If
    
    
    If ColIndex = 5 Then
        If tdbgFlujo.Columns(2) = "M" Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub tdbgFlujo_ButtonClick(ByVal ColIndex As Integer)
    Dim Tipo As String
    Tipo = CE(tdbgFlujo.Columns(2).Value)
    
    If ColIndex = 5 And Tipo <> "M" And Tipo <> "" Then
        
        frmFormulas.pFormula = CE(tdbgFlujo.Columns(4).Value)
        frmFormulas.pObservacion = CE(tdbgFlujo.Columns(5).Value)
        frmFormulas.pFormulario = Me.Name
        frmFormulas.pTipo = Tipo
        frmFormulas.pMetodo = cboMetodo.Text
        frmFormulas.pPeriodo = tdbcMes.BoundText
        frmFormulas.Show vbModal
    End If
End Sub

Private Sub tdbgFlujo_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    On Error GoTo serror
    
    If lArrFlujo(Bookmark, 6) = "S" Or lArrFlujo(Bookmark, 6) = "T" Then
        RowStyle.BackColor = gsColorDesactProv
    End If
    
    Exit Sub
serror:
End Sub

Public Sub UpdateGrilla()
    On Error Resume Next
    DoEvents
    tdbgFlujo.Update
    DoEvents
End Sub

Private Sub tdbgFlujo_KeyDown(KeyCode As Integer, Shift As Integer)
    If gsColumna = 5 Then
       If KeyCode <> 13 And KeyCode <> vbKeyEnd And KeyCode <> vbKeyHome And _
          KeyCode <> 46 And KeyCode <> vbKeyF2 And KeyCode <> vbKeyBack And _
          KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyUp And _
          KeyCode <> vbKeyDown Then
            If Not IsNumeric(Chr(KeyCode)) Then
               KeyCode = 0
               Exit Sub
            End If
       End If
    End If

    If Mid(gsGrupo, 3, 1) = "1" Or gsGrupo = gsPrivilegioAdmin Then
    
        If KeyCode = vbKeyDelete And gsColumna = 5 Then
            tdbgFlujo.Columns(4) = ""
            tdbgFlujo.Columns(5) = ""
            UpdateGrilla
        End If
        
        If KeyCode = vbKeyDelete And gsColumna = 3 Then
            tdbgFlujo.Columns(2) = ""
            tdbgFlujo.Columns(3) = ""
            tdbgFlujo.Columns(4) = ""
            tdbgFlujo.Columns(5) = ""
            UpdateGrilla
        End If
    
    End If
End Sub

Private Sub tdbgFlujo_KeyPress(KeyAscii As Integer)
    If gsColumna = 5 Then
       If KeyAscii <> 13 And KeyAscii <> vbKeyEnd And KeyAscii <> vbKeyHome And _
          KeyAscii <> 46 And KeyAscii <> vbKeyF2 And KeyAscii <> vbKeyBack And _
          KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyUp And _
          KeyAscii <> vbKeyDown Then
            If Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Exit Sub
            End If
       End If
    End If

    If KeyAscii = 13 And gsColumna >= 4 Then
       On Error GoTo serror
       UpdateGrilla
    
       tdbgFlujo.Col = 3
       tdbgFlujo.Bookmark = tdbgFlujo.Bookmark + 1
serror:
       pSetFocus tdbgFlujo
    
    End If
End Sub

Private Sub tdbgFlujo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gsColumna = tdbgFlujo.Col
    On Error GoTo serror
    If lArrFlujo(tdbgFlujo.Bookmark, 7) = "S" Then
        cmdEliminaItem.Enabled = False
    Else
        cmdEliminaItem.Enabled = True
    End If
    Exit Sub
serror:
End Sub

