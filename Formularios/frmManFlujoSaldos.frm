VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManFlujoSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Saldos iniciales del Flujo de Efectivo"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "frmManFlujoSaldos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   11460
   Begin VB.Frame Frame1 
      Caption         =   " Ejercicio y Periodo de los saldos iniciales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   11175
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   7200
         TabIndex        =   6
         Top             =   270
         Width           =   3390
         _ExtentX        =   5980
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
         _PropDict       =   $"frmManFlujoSaldos.frx":0ECA
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
      Begin TrueOleDBList70.TDBCombo tdbcAnio 
         Height          =   300
         Left            =   1530
         TabIndex        =   7
         Top             =   270
         Width           =   3390
         _ExtentX        =   5980
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
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).DividerStyle=   2
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
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
         _PropDict       =   $"frmManFlujoSaldos.frx":0F51
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
         Caption         =   "Ejercicio"
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
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   735
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
         Index           =   1
         Left            =   6345
         TabIndex        =   8
         Top             =   315
         Width           =   660
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgFlujo 
      Height          =   4425
      Left            =   90
      TabIndex        =   0
      Top             =   1485
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   7805
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
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
      Columns(2).Caption=   "Saldo Inicial"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "External Editor"
      Columns(2).ExternalEditor=   "TDBNumberPorc"
      Columns(2).ExternalEditor.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Titulo"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=139776"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=13573"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=13494"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=139780"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1).AutoDropDown=1"
      Splits(0)._ColumnProps(14)=   "Column(1).DropDownList=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=661"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=582"
      Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=139778"
      Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=139780"
      Splits(0)._ColumnProps(28)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(29)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).Size  =   3
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   688
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   14215660
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=4"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1984"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=131584"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=13573"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=13494"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=139780"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(1).AutoDropDown=1"
      Splits(1)._ColumnProps(18)=   "Column(1).DropDownList=1"
      Splits(1)._ColumnProps(19)=   "Column(2).Width=661"
      Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=582"
      Splits(1)._ColumnProps(22)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(23)=   "Column(2)._ColStyle=131586"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=2725"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=2646"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=131588"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
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
      CellTipsWidth   =   0
      DeadAreaBackColor=   16777215
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   2
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
      _StyleDefs(25)  =   "Splits(0).Style:id=79,.parent=1"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=79,.alignment=0,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=80"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=81"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=83"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=79,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=80"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=81"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=83"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=62,.parent=79,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(46)  =   ":id=62,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=80"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=81"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=83"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=16,.parent=79,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=80"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=81"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=83"
      _StyleDefs(54)  =   "Splits(1).Style:id=17,.parent=1"
      _StyleDefs(55)  =   "Splits(1).CaptionStyle:id=26,.parent=4"
      _StyleDefs(56)  =   "Splits(1).HeadingStyle:id=18,.parent=2"
      _StyleDefs(57)  =   "Splits(1).FooterStyle:id=19,.parent=3"
      _StyleDefs(58)  =   "Splits(1).InactiveStyle:id=20,.parent=5"
      _StyleDefs(59)  =   "Splits(1).SelectedStyle:id=22,.parent=6"
      _StyleDefs(60)  =   "Splits(1).EditorStyle:id=21,.parent=7"
      _StyleDefs(61)  =   "Splits(1).HighlightRowStyle:id=23,.parent=8"
      _StyleDefs(62)  =   "Splits(1).EvenRowStyle:id=24,.parent=9"
      _StyleDefs(63)  =   "Splits(1).OddRowStyle:id=25,.parent=10"
      _StyleDefs(64)  =   "Splits(1).RecordSelectorStyle:id=27,.parent=11"
      _StyleDefs(65)  =   "Splits(1).FilterBarStyle:id=28,.parent=12"
      _StyleDefs(66)  =   "Splits(1).Columns(0).Style:id=46,.parent=17,.alignment=0,.locked=0"
      _StyleDefs(67)  =   "Splits(1).Columns(0).HeadingStyle:id=43,.parent=18"
      _StyleDefs(68)  =   "Splits(1).Columns(0).FooterStyle:id=44,.parent=19"
      _StyleDefs(69)  =   "Splits(1).Columns(0).EditorStyle:id=45,.parent=21"
      _StyleDefs(70)  =   "Splits(1).Columns(1).Style:id=54,.parent=17,.locked=-1"
      _StyleDefs(71)  =   "Splits(1).Columns(1).HeadingStyle:id=51,.parent=18"
      _StyleDefs(72)  =   "Splits(1).Columns(1).FooterStyle:id=52,.parent=19"
      _StyleDefs(73)  =   "Splits(1).Columns(1).EditorStyle:id=53,.parent=21"
      _StyleDefs(74)  =   "Splits(1).Columns(2).Style:id=58,.parent=17,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(75)  =   "Splits(1).Columns(2).HeadingStyle:id=55,.parent=18"
      _StyleDefs(76)  =   "Splits(1).Columns(2).FooterStyle:id=56,.parent=19"
      _StyleDefs(77)  =   "Splits(1).Columns(2).EditorStyle:id=57,.parent=21"
      _StyleDefs(78)  =   "Splits(1).Columns(3).Style:id=66,.parent=17"
      _StyleDefs(79)  =   "Splits(1).Columns(3).HeadingStyle:id=63,.parent=18"
      _StyleDefs(80)  =   "Splits(1).Columns(3).FooterStyle:id=64,.parent=19"
      _StyleDefs(81)  =   "Splits(1).Columns(3).EditorStyle:id=65,.parent=21"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumberPorc 
      Height          =   285
      Left            =   10305
      TabIndex        =   4
      Top             =   1845
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManFlujoSaldos.frx":0FD8
      Caption         =   "frmManFlujoSaldos.frx":0FF8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManFlujoSaldos.frx":105C
      Keys            =   "frmManFlujoSaldos.frx":107A
      Spin            =   "frmManFlujoSaldos.frx":10B4
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
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
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
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   7470
      TabIndex        =   10
      Top             =   1035
      Width           =   1575
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoSaldos.frx":10DC
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   5805
      TabIndex        =   3
      ToolTipText     =   " Eliminar todas las cuentas de la lista "
      Top             =   1035
      Width           =   1575
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoSaldos.frx":1676
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      ToolTipText     =   " Graba la lista mostrada "
      Top             =   1035
      Width           =   1575
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoSaldos.frx":1C10
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   375
      Left            =   2475
      TabIndex        =   1
      ToolTipText     =   " Vuelve a cargar los datos almacenados "
      Top             =   1035
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmManFlujoSaldos.frx":21AA
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmManFlujoSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lArrFlujo As New XArrayDB
Dim lArrDetalle(6) As Variant
Dim gsGrupo As String
Dim gsColumna As Integer
Const NUM_COL = 4

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdEliminarTodo_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If

    If MsgBox("Deseas eliminar todas las cuentas de la lista", vbYesNo + vbQuestion) = vbYes Then
        Dim i As Integer
        For i = 0 To lArrFlujo.Count(1) - 1
            If lArrFlujo(i, 3) = "S" Then
                lArrFlujo(i, 2) = ""
            Else
                lArrFlujo(i, 2) = 0
            End If
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
    lArrDetalle(3) = tdbcAnio.BoundText
    lArrDetalle(4) = tdbcMes.BoundText
    lArrDetalle(5) = CE(lArrFlujo(item, 0))
    lArrDetalle(6) = NE(lArrFlujo(item, 2))
End Function

Private Sub Grabar()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas
    

    Dim lArrDet(4) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcAnio.BoundText
    lArrDet(4) = tdbcMes.BoundText

    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoSaldos", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If NE(lArrFlujo(i, 2)) <> 0 And CE(lArrFlujo(i, 0)) <> "XX" Then
            If CargaArregloDet(i) = True Then
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoSaldos", lArrDetalle(), False) = False Then
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

Private Sub cmdImportar_Click()
    If MsgBox("Desea importar los datos del mes seleccionado", vbYesNo + vbQuestion) = vbYes Then
        cmdImportar.Enabled = False
        Screen.MousePointer = vbHourglass
        
        GeneraArreglo tdbcAnioImportar.BoundText, tdbcMesImportar.BoundText
        DoEvents
        cmdImportar.Enabled = True
        Screen.MousePointer = vbNormal
        
        tdbgFlujo.ReBind
        On Error Resume Next
        tdbgFlujo.Row = 0
        tdbgFlujo.Bookmark = 0
    End If
End Sub


Private Sub GeneraArreglo(Anio As String, Mes As String)
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion
    
    If tdbcMes.Text <> "" Then
        sql = "spCn_FlujoSaldos 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "','" & Anio & "', '" & Mes & "'"
        Call GridArreglo(lArrFlujo, tdbgFlujo, sql)
        
        If lArrFlujo.Count(2) < NUM_COL Then
           lArrFlujo.ReDim 0, Filas, 0, NUM_COL
        End If
    Else
        lArrFlujo.ReDim 0, 0, 1, NUM_COL
        'Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        
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


    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass

    DoEvents
    
    GeneraArreglo tdbcAnio.BoundText, tdbcMes.BoundText
    
    DoEvents
    SumarTotal
    
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
    
    tdbgFlujo.ReBind
    On Error Resume Next
    tdbgFlujo.Row = 0
    tdbgFlujo.Bookmark = 0
    
    
    
End Sub

Private Sub cmdsalir_Click()
Unload Me
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

Private Sub LlenaCombos()
    Dim sql As String
    Dim i As Double
    
    tdbcAnio.Clear

    
    For i = 2000 To NE(gsAnio) - 1
        tdbcAnio.AddItem CE(i) & ";" & CE(i)
    Next i
    

    tdbcAnio.Bookmark = 0
    tdbcAnio.ListField = "column1"
    tdbcAnio.BoundColumn = "column1"
    tdbcAnio.BoundText = gsAnio - 1
    tdbcAnio.ReBind
    
    
    
    '---------------------------------------------------------
    Call LlenaComboMesApeAddItem(tdbcMes)
    DoEvents
    tdbcMes.BoundText = gsPeriodo
    '---------------------------------------------------------
End Sub


Private Sub SumarTotal()
    If lArrFlujo Is Nothing Or IsNull(tdbgFlujo.Bookmark) Then
        Exit Sub
    End If

    Dim iFila As Integer
    
    
    Dim s_debe As Double
    Dim i As Integer
    
    On Error Resume Next
    
    s_debe = 0
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If lArrFlujo.Value(i, 0) <> "XX" Then
            s_debe = s_debe + NE(lArrFlujo.Value(i, 2))
        Else
            lArrFlujo.Value(i, 2) = s_debe
            
            s_debe = 0
        End If
    Next
    
    tdbgFlujo.Refresh
End Sub
Private Sub Form_Load()
    Me.Height = 6525
    Me.Width = 11550
    
    LlenaCombos
    
    tdbgFlujo.Splits(0).MarqueeStyle = dbgHighlightRow
    tdbgFlujo.Splits(1).MarqueeStyle = dbgFloatingEditor
    DoEvents
    
    tdbgFlujo.FetchRowStyle = True
  
    cmdRefresh_Click
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabar.Enabled = False
        cmdEliminarTodo.Enabled = False
'        cmdImportar.Enabled = False
        tdbgFlujo.Splits(0).Locked = True
    Else
        cmdGrabar.Enabled = True
        cmdEliminarTodo.Enabled = True
        tdbgFlujo.Splits(0).Locked = False
    End If
    
End Sub

Private Sub Form_Resize()
    tdbgFlujo.Width = Me.Width - 200
    tdbgFlujo.Height = Me.Height - tdbgFlujo.Top - 500
End Sub

Private Sub tdbcMes_ItemChange()
    cmdRefresh_Click
    DoEvents
    pSetFocus tdbcMes
End Sub


Private Sub tdbgFlujo_AfterColEdit(ByVal ColIndex As Integer)
    UpdateGrilla
    
    SumarTotal
End Sub

Private Sub tdbgFlujo_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If lArrFlujo Is Nothing Or IsNull(tdbgFlujo.Bookmark) Then
        Exit Sub
    End If
    If lArrFlujo(tdbgFlujo.Bookmark, 0) = "XX" Then
        If ColIndex = 2 Then
            Cancel = 1
        End If
    End If
    If lArrFlujo.Count(2) >= NUM_COL Then
        If ColIndex <> 1 And CE(lArrFlujo(tdbgFlujo.Bookmark, 3)) = "S" Then
           Cancel = 1
        End If
    End If
End Sub

Private Sub tdbgFlujo_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    'If lArrFlujo Is Nothing Or IsNull(tdbgFlujo.Bookmark) Then
    '    Exit Sub
    'End If
    
    On Error GoTo serror
    
    If lArrFlujo(Bookmark, 0) = "XX" Then
        RowStyle.BackColor = gsColorDesactProv
    End If
    
    If Split = 0 Then
    If lArrFlujo(Bookmark, 0) <> "XX" And lArrFlujo(Bookmark, 4) = "G1" Then
        RowStyle.BackColor = &HFFF9D7   'celeste
    End If
    
    If lArrFlujo(Bookmark, 0) <> "XX" And lArrFlujo(Bookmark, 4) = "G2" Then
        RowStyle.BackColor = &HCEFFF8   'amarillo
    End If
    End If
    
    Exit Sub
serror:
End Sub

Private Sub UpdateGrilla()
    On Error Resume Next
    tdbgFlujo.Update
End Sub



Private Sub tdbgFlujo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gsColumna = tdbgFlujo.Col
    
    On Error GoTo serror
    If lArrFlujo(tdbgFlujo.Bookmark, 0) = "XX" Then
        TDBNumberPorc.BackColor = gsColorDesactProv
        
    Else
        TDBNumberPorc.BackColor = gsColorActivado
        
    End If
serror:
End Sub

Private Sub TDBNumberPorc_InvalidInput()
    TDBNumberPorc.Value = 0
End Sub

Private Sub TDBNumberPorc_InvalidRange(Restore As Boolean)
    TDBNumberPorc.Value = 0
End Sub


