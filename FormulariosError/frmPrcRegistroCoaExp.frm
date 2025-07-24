VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcRegistroCoaExp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Datos al COA Exportación"
   ClientHeight    =   6030
   ClientLeft      =   1200
   ClientTop       =   3570
   ClientWidth     =   11550
   Icon            =   "frmPrcRegistroCoaExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   11550
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   300
      Left            =   5085
      TabIndex        =   7
      Tag             =   "enabled"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   529
      Calendar        =   "frmPrcRegistroCoaExp.frx":1982
      Caption         =   "frmPrcRegistroCoaExp.frx":1A84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcRegistroCoaExp.frx":1AE8
      Keys            =   "frmPrcRegistroCoaExp.frx":1B06
      Spin            =   "frmPrcRegistroCoaExp.frx":1B72
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
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   375
      Left            =   3585
      TabIndex        =   3
      Top             =   3225
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      Calculator      =   "frmPrcRegistroCoaExp.frx":1B9A
      Caption         =   "frmPrcRegistroCoaExp.frx":1BBA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcRegistroCoaExp.frx":1C26
      Keys            =   "frmPrcRegistroCoaExp.frx":1C44
      Spin            =   "frmPrcRegistroCoaExp.frx":1C8E
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1802698757
      MinValueVT      =   1769209861
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   855
      TabIndex        =   0
      Top             =   315
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
      _PropDict       =   $"frmPrcRegistroCoaExp.frx":1CB6
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
   Begin TrueOleDBList70.TDBCombo tdbcMoneda 
      Height          =   300
      Left            =   9450
      TabIndex        =   1
      Tag             =   "_"
      Top             =   675
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      _PropDict       =   $"frmPrcRegistroCoaExp.frx":1D3D
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
   Begin TrueOleDBList70.TDBCombo tdbcLibro 
      Height          =   300
      Left            =   855
      TabIndex        =   8
      Tag             =   "enabled"
      Top             =   675
      Width           =   3420
      _ExtentX        =   6033
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
      _PropDict       =   $"frmPrcRegistroCoaExp.frx":1DC4
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4830
      Left            =   45
      TabIndex        =   2
      Top             =   1125
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   8520
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tipo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Codigo"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Ruc"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Razon Social"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TD"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Serie"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Numero"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fec Emision"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "External Editor"
      Columns(7).ExternalEditor=   "TDBDate1"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Tipo"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Monto"
      Columns(9).DataField=   "Soles"
      Columns(9).NumberFormat=   "External Editor"
      Columns(9).ExternalEditor=   "TDBNumber1"
      Columns(9).ExternalEditor.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Tc"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "External Editor"
      Columns(10).ExternalEditor=   "TDBNumber1"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Mon Ext."
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "External Editor"
      Columns(11).ExternalEditor=   "TDBNumber1"
      Columns(11).ExternalEditor.vt=   8
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Valor FOB"
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "External Editor"
      Columns(12).ExternalEditor=   "TDBNumber1"
      Columns(12).ExternalEditor.vt=   8
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Val. Flete"
      Columns(13).DataField=   ""
      Columns(13).NumberFormat=   "External Editor"
      Columns(13).ExternalEditor=   "TDBNumber1"
      Columns(13).ExternalEditor.vt=   8
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Num DUE"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Fec Numer."
      Columns(15).DataField=   ""
      Columns(15).NumberFormat=   "External Editor"
      Columns(15).ExternalEditor=   "TDBDate1"
      Columns(15).ExternalEditor.vt=   8
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Fec Embarq."
      Columns(16).DataField=   ""
      Columns(16).NumberFormat=   "External Editor"
      Columns(16).ExternalEditor=   "TDBDate1"
      Columns(16).ExternalEditor.vt=   8
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Glosa"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Interno"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Empresa"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Año"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Periodo"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Libro"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Voucher"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Item"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   25
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=25"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=503"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=423"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=926"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=847"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2223"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2143"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=3387"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=3307"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=529"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=450"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1138"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1058"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1746"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2355"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=688"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=609"
      Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=2170"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2090"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(10).Width=1349"
      Splits(0)._ColumnProps(58)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(10)._WidthInPix=1270"
      Splits(0)._ColumnProps(60)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(61)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(62)=   "Column(11).Width=2223"
      Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=2143"
      Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(67)=   "Column(12).Width=1984"
      Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=1905"
      Splits(0)._ColumnProps(70)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(71)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(72)=   "Column(13).Width=1879"
      Splits(0)._ColumnProps(73)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(13)._WidthInPix=1799"
      Splits(0)._ColumnProps(75)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(76)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(77)=   "Column(14).Width=2249"
      Splits(0)._ColumnProps(78)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(14)._WidthInPix=2170"
      Splits(0)._ColumnProps(80)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(81)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(82)=   "Column(15).Width=2302"
      Splits(0)._ColumnProps(83)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(15)._WidthInPix=2223"
      Splits(0)._ColumnProps(85)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(86)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(87)=   "Column(16).Width=2328"
      Splits(0)._ColumnProps(88)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(16)._WidthInPix=2249"
      Splits(0)._ColumnProps(90)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(91)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(92)=   "Column(17).Width=6826"
      Splits(0)._ColumnProps(93)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(17)._WidthInPix=6747"
      Splits(0)._ColumnProps(95)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(96)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(97)=   "Column(18).Width=344"
      Splits(0)._ColumnProps(98)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(18)._WidthInPix=265"
      Splits(0)._ColumnProps(100)=   "Column(18).AllowSizing=0"
      Splits(0)._ColumnProps(101)=   "Column(18)._ColStyle=8708"
      Splits(0)._ColumnProps(102)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(18).AllowFocus=0"
      Splits(0)._ColumnProps(104)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(105)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(106)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(108)=   "Column(19).AllowSizing=0"
      Splits(0)._ColumnProps(109)=   "Column(19)._ColStyle=8708"
      Splits(0)._ColumnProps(110)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(111)=   "Column(19).AllowFocus=0"
      Splits(0)._ColumnProps(112)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(113)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(20).AllowSizing=0"
      Splits(0)._ColumnProps(117)=   "Column(20)._ColStyle=8708"
      Splits(0)._ColumnProps(118)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(119)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(120)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(121)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(125)=   "Column(21)._ColStyle=8708"
      Splits(0)._ColumnProps(126)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(127)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(128)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(129)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(22).AllowSizing=0"
      Splits(0)._ColumnProps(133)=   "Column(22)._ColStyle=8708"
      Splits(0)._ColumnProps(134)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(135)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(136)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(137)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(138)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(140)=   "Column(23).AllowSizing=0"
      Splits(0)._ColumnProps(141)=   "Column(23)._ColStyle=8708"
      Splits(0)._ColumnProps(142)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(143)=   "Column(23).AllowFocus=0"
      Splits(0)._ColumnProps(144)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(145)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(146)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(148)=   "Column(24).AllowSizing=0"
      Splits(0)._ColumnProps(149)=   "Column(24)._ColStyle=8708"
      Splits(0)._ColumnProps(150)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(151)=   "Column(24).AllowFocus=0"
      Splits(0)._ColumnProps(152)=   "Column(24).Order=25"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.bgcolor=&HFFFFFF&,.locked=-1"
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
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=110,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=102,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=28,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=32,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=82,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=90,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=87,.parent=14"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=88,.parent=15"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=89,.parent=17"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=86,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=14"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=15"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=17"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=106,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=118,.parent=13,.locked=-1"
      _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=115,.parent=14"
      _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=116,.parent=15"
      _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=117,.parent=17"
      _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=122,.parent=13,.locked=-1"
      _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=119,.parent=14"
      _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=120,.parent=15"
      _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=121,.parent=17"
      _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=126,.parent=13,.locked=-1"
      _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=123,.parent=14"
      _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=124,.parent=15"
      _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=125,.parent=17"
      _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=130,.parent=13,.locked=-1"
      _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=127,.parent=14"
      _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=128,.parent=15"
      _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=129,.parent=17"
      _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=142,.parent=13,.locked=-1"
      _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=139,.parent=14"
      _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=140,.parent=15"
      _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=141,.parent=17"
      _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=138,.parent=13,.locked=-1"
      _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=135,.parent=14"
      _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=136,.parent=15"
      _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=137,.parent=17"
      _StyleDefs(133) =   "Splits(0).Columns(24).Style:id=134,.parent=13,.locked=-1"
      _StyleDefs(134) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=14"
      _StyleDefs(135) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=15"
      _StyleDefs(136) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=17"
      _StyleDefs(137) =   "Named:id=33:Normal"
      _StyleDefs(138) =   ":id=33,.parent=0"
      _StyleDefs(139) =   "Named:id=34:Heading"
      _StyleDefs(140) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(141) =   ":id=34,.wraptext=-1"
      _StyleDefs(142) =   "Named:id=35:Footing"
      _StyleDefs(143) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(144) =   "Named:id=36:Selected"
      _StyleDefs(145) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(146) =   "Named:id=37:Caption"
      _StyleDefs(147) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(148) =   "Named:id=38:HighlightRow"
      _StyleDefs(149) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(150) =   "Named:id=39:EvenRow"
      _StyleDefs(151) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(152) =   "Named:id=40:OddRow"
      _StyleDefs(153) =   ":id=40,.parent=33"
      _StyleDefs(154) =   "Named:id=41:RecordSelector"
      _StyleDefs(155) =   ":id=41,.parent=34"
      _StyleDefs(156) =   "Named:id=42:FilterBar"
      _StyleDefs(157) =   ":id=42,.parent=33"
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   9405
      TabIndex        =   15
      Top             =   270
      Width           =   1575
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":1E4B
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdListar 
      Height          =   375
      Left            =   4410
      TabIndex        =   14
      ToolTipText     =   "Cargar nueva Configuración"
      Top             =   270
      Width           =   1575
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":23E5
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   4410
      TabIndex        =   13
      ToolTipText     =   "Grabar modificaciones"
      Top             =   675
      Width           =   1575
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":297F
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   7740
      TabIndex        =   12
      ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
      Top             =   675
      Width           =   1575
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":2F19
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   7740
      TabIndex        =   11
      ToolTipText     =   "Eliminar el movimientos seleccionado"
      Top             =   270
      Width           =   1575
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":34B3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdTodos 
      Height          =   375
      Left            =   6075
      TabIndex        =   10
      ToolTipText     =   "Insertar todos los movimientos del libro y mes seleccionado"
      Top             =   675
      Width           =   1575
      Caption         =   " Insertar Todo"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":3A4D
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   6075
      TabIndex        =   9
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   270
      Width           =   1575
      Caption         =   " Insertar Item"
      PicturePosition =   327683
      Size            =   "2778;661"
      Picture         =   "frmPrcRegistroCoaExp.frx":3FE7
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
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
      Left            =   9450
      TabIndex        =   6
      Top             =   675
      Visible         =   0   'False
      Width           =   660
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
      TabIndex        =   5
      Top             =   360
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
      TabIndex        =   4
      Top             =   720
      Width           =   420
   End
End
Attribute VB_Name = "frmPrcRegistroCoaExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrDatos As New XArrayDB
Dim lControl As String
Dim lArrDet() As Variant
Dim Sw As Boolean
Dim lsLibroCom As String
Dim lsLibroVen As String
Dim NUM_FILAS As Integer
Dim NUM_COLUMNAS As Integer

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub EliminaTodo()
        Dim clsMante As New clsMantoTablas
        
        Call EliminaArreglo
        
        clsMante.InicializaClase
        clsMante.BeginTrans
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoCoa", lArrDet(), False) = False Then
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

Private Sub EliminaArreglo()
        'lArrDatos.ReDim 0, 0, 0, 30
        ReDim lArrDet(8)
        
        lArrDet(0) = "ELIMINAR"
        lArrDet(1) = ""    ' *** @Ase_cNummov
        lArrDet(2) = gsEmpresa    ' *** @Emp_cCodigo
        lArrDet(3) = gsAnio    ' *** @Pan_cAnio
        lArrDet(4) = Me.tdbcMes.BoundText    ' *** @Per_cPeriodo
        'lArrDet(5) = lArrDatos(item, 24)    ' *** @Lib_cTipoLibro
        'lArrDet(6) = lArrDatos(item, 25)    ' *** @Ase_nVoucher
        lArrDet(7) = 0                   ' *** @Dco_nItem
        lArrDet(8) = "E"                    ' *** @Dco_cTipo
End Sub

Private Sub cmdEliminarTodo_Click()
    cmdEliminarTodo.Enabled = False
    DoEvents
    
    If MsgBox("Deseas eliminar todos los registros de Exportación", vbYesNo + vbInformation) = vbYes Then
        EliminaTodo
        llenaGrilla
    End If
    
    DoEvents
    cmdEliminarTodo.Enabled = True
End Sub

Private Sub cmdEliminaItem_Click()
    If CE(TDBGrid1.Columns(TDBGrid1.Bookmark)) <> "" Then
        cmdEliminaItem.Enabled = False
        DoEvents
        
        lArrDatos.DeleteRows (Me.TDBGrid1.Bookmark)
'        Grabar
        
        DoEvents
        cmdEliminaItem.Enabled = True
        
    End If
    TDBGrid1.ReBind
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
    Dim i As Integer
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    On Error GoTo ERROR
    Grabar = True
    
    Dim Fila As Integer

    Fila = CuentaFilas
    
    DoEvents
    EliminaTodo
    DoEvents
    
    clsMante.InicializaClase
    clsMante.BeginTrans
    
    On Error GoTo ERROR
        For i = 0 To Fila - 1
            If CE(lArrDatos(i, 1)) <> "" Then
                Call CargaArregloDet(i)
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoCoa", lArrDet(), False) = False Then
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
    

    Exit Function
ERROR:
    Grabar = False
    Screen.MousePointer = vbNormal


End Function


Private Sub cmdGrabar_Click()
    cmdGrabar.Enabled = False
    DoEvents
    If Grabar = True Then

        Mensajes "Datos se grabaron con exito.", vbInformation
        
    End If
    DoEvents
    cmdGrabar.Enabled = True
End Sub

Private Sub CargaArregloDet(item As Integer)
    On Error Resume Next
    Dim i As Integer
    i = 0
    
'    If lArrDetalle(num, 22) = "  /  /    " Then lArrDetalle(num, 22) = Null
'    If lArrDetalle(num, 22) = "" Then lArrDetalle(num, 22) = Null
    ReDim lArrDet(40) As Variant
    If lArrDatos.Count(1) = 0 Then
        'lArrDatos.ReDim 0, 0, 0, 30
        lArrDet(0) = "ELIMINAR"    ' *** @Ase_cNummov
        lArrDet(1) = ""    ' *** @Ase_cNummov
        lArrDet(2) = gsEmpresa    ' *** @Emp_cCodigo
        lArrDet(3) = gsAnio    ' *** @Pan_cAnio
        lArrDet(4) = Me.tdbcMes.BoundText    ' *** @Per_cPeriodo
        'lArrDet(5) = lArrDatos(item, 24)    ' *** @Lib_cTipoLibro
        'lArrDet(6) = lArrDatos(item, 25)    ' *** @Ase_nVoucher
        lArrDet(7) = 0                   ' *** @Dco_nItem
        lArrDet(8) = "E"                    ' *** @Dco_cTipo
        Exit Sub
    End If
    lArrDet(0) = "INSERTAR"             ' *** @Accion
    lArrDet(1) = lArrDatos(item, 18 + i)  ' *** @Ase_cNummov
    lArrDet(2) = CE(lArrDatos(item, 19 + i))  ' *** @Emp_cCodigo
    lArrDet(3) = CE(lArrDatos(item, 20 + i))  ' *** @Pan_cAnio
    lArrDet(4) = CE(lArrDatos(item, 21 + i))  ' *** @Per_cPeriodo
    lArrDet(5) = CE(lArrDatos(item, 22 + i))  ' *** @Lib_cTipoLibro
    lArrDet(6) = lArrDatos(item, 23 + i)  ' *** @Ase_nVoucher
    lArrDet(7) = item                   ' *** @Dco_nItem
    lArrDet(8) = "E"                    ' *** @Dco_cTipo
    lArrDet(9) = CE(lArrDatos(item, 8 + i))   ' *** @Dco_cTipoBS
    lArrDet(10) = CE(lArrDatos(item, 17 + i)) ' *** @Dco_cGlosa
    lArrDet(11) = CE(tdbcMoneda.BoundText)  ' *** @Dco_cMoneda
    lArrDet(12) = CE(lArrDatos(item, 10 + i)) ' *** @Asd_nTipoCambio
    lArrDet(13) = NE(lArrDatos(item, 9 + i)) ' *** @Dco_nMonto
    lArrDet(14) = CE(lArrDatos(item, 4 + i))  ' *** @Dco_cTipoDoc
    lArrDet(15) = CE(lArrDatos(item, 5 + i))  ' *** @Dco_cSerieDoc
    lArrDet(16) = CE(lArrDatos(item, 6 + i))  ' *** @Dco_cNumDoc
    
    lArrDet(17) = CE(lArrDatos(item, 7 + i))  ' *** @Dco_dFecDoc
    
    If CE(lArrDet(17)) = "" Then lArrDet(17) = Null
    
    lArrDet(18) = CE(lArrDatos(item, 0 + i))  ' *** @Dco_cTipoEntidad
    lArrDet(19) = CE(lArrDatos(item, 1 + i))  ' *** @Dco_cCodEntidad
    
    lArrDet(20) = CE(lArrDatos(item, 2 + i))  ' *** @Dco_cNumRuc
    lArrDet(21) = CE(lArrDatos(item, 3 + i))  ' *** @Dco_cRazonSocial
    lArrDet(22) = ""                    ' *** @Dco_cTipoIGV
    lArrDet(23) = Null                  ' *** @Dco_dFecDocRef
    lArrDet(24) = ""                    ' *** @Dco_cSerieDocRef
    lArrDet(25) = ""                    ' *** @Dco_cNumDocRef
    lArrDet(26) = Null                  ' *** @Dco_dFechaEmision
    lArrDet(27) = Null                  ' *** @Dco_cFechaPago
    lArrDet(28) = CE(lArrDatos(item, 14 + i)) ' *** @Dco_nNumDue
    lArrDet(29) = 0                     ' *** @Dco_nMontoCIF
    
    lArrDet(30) = 0                     ' *** @Dco_nMontoAdvalorem
    lArrDet(31) = 0                     ' *** @Dco_nMontoInafecto
    lArrDet(32) = 0                     ' *** @Dco_nMontoIGV
    lArrDet(33) = 0                     ' *** @Dco_nMontoIPM
    lArrDet(34) = 0                     ' *** @Dco_nMontoISC
    lArrDet(35) = NE(lArrDatos(item, 12 + i)) ' *** @Dco_nValorFOB
    lArrDet(36) = NE(lArrDatos(item, 13 + i)) ' *** @Dco_nValorFlete
    
    If lArrDatos(item, 15) = "" Then
        lArrDet(37) = Null
    Else
        lArrDet(37) = CE(lArrDatos(item, 15 + i)) ' *** @Dco_dFechaNumera
        If CE(lArrDet(37)) = "" Then lArrDet(37) = Null
    End If
    If lArrDatos(item, 16) = "" Then
        lArrDet(38) = Null
    Else
        lArrDet(38) = CE(lArrDatos(item, 16 + i)) ' *** @Dco_dFechaEmbarque
        If CE(lArrDet(38)) = "" Then lArrDet(38) = Null
    End If
    
    lArrDet(39) = ""                    ' *** @Dco_cEstado
    lArrDet(40) = gsUsuario             ' *** @Dco_cUserCrea
    
End Sub

Private Sub cmdInsertarItem_Click()
    cmdInsertarItem.Enabled = False
    DoEvents
    gsPeriodoCOA = Me.tdbcMes.BoundText
    Call LlamaBuscar(frmBusCoa, "Provisiones", lControl, "Provisiones", Me, gsPeriodo, tdbcLibro.BoundText)
    DoEvents
    cmdInsertarItem.Enabled = True
End Sub

Private Sub cmdListar_Click()
    cmdListar.Enabled = False
    DoEvents
    llenaGrilla
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

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTodos_Click()
    ' *** Jalar todos los datos dependiendo del Tipo de Libro
    cmdTodos.Enabled = False
    DoEvents
    
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim i As Integer
    Dim lrsProvision As New ADODB.Recordset
    
    Set clDatos = New clsMantoTablas
    Set lrsProvision = New ADODB.Recordset
    sqlSp = "spCn_ConsultaProvisionesCoa 'SEL_TODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcLibro.BoundText & "' "
    arrDatos = Array(sqlSp)
    Set lrsProvision = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsProvision.State <> 0 Then
        ' *** Cargar los datos de la grilla
        Screen.MousePointer = vbHourglass
        
        i = BuscaCelda
        
        Do While Not lrsProvision.EOF
            lArrDatos(i, 0) = CE(lrsProvision("Ten_cTipoEntidad").Value) ' *** Tipo
            lArrDatos(i, 1) = CE(lrsProvision("Ent_cCodEntidad").Value) ' *** Codigo
            lArrDatos(i, 2) = CE(lrsProvision("Ent_nRuc").Value)        ' *** Ruc
            lArrDatos(i, 3) = CE(lrsProvision("Ent_cPersona").Value)    ' *** Razon Social
            lArrDatos(i, 4) = CE(lrsProvision("Asd_cTipoDoc").Value)    ' *** Td
            lArrDatos(i, 5) = CE(lrsProvision("Asd_cSerieDoc").Value)   ' *** Serie
            lArrDatos(i, 6) = CE(lrsProvision("Asd_cNumDoc").Value)     ' *** Numero
            lArrDatos(i, 7) = CE(lrsProvision("Asd_dFecDoc").Value)    ' *** Fecha
            lArrDatos(i, 8) = ""                                    ' *** Tipo Exp
            lArrDatos(i, 9) = NE(lrsProvision("Soles").Value) ' *** Soles
            lArrDatos(i, 10) = NE(lrsProvision("Asd_nTipoCambio").Value) ' *** Tc
            
            lArrDatos(i, 11) = CE(lrsProvision("Dolares").Value) ' *** Dolares
            lArrDatos(i, 17) = CE(lrsProvision("asd_cGlosa").Value)    ' *** Glosa
            lArrDatos(i, 18) = CE(lrsProvision("Ase_cNummov").Value) ' *** Interno
            lArrDatos(i, 19) = CE(lrsProvision("Emp_cCodigo").Value) ' *** Empresa
            lArrDatos(i, 20) = CE(lrsProvision("Pan_cAnio").Value) ' *** Año
            lArrDatos(i, 21) = CE(lrsProvision("Per_cPeriodo").Value) ' *** Periodo
            lArrDatos(i, 22) = CE(lrsProvision("Lib_cTipoLibro").Value) ' *** Libro
            lArrDatos(i, 23) = CE(lrsProvision("Ase_nVoucher").Value) ' *** Voucher

            ' *** Aqui llamar a un sp q busq el valor de los datods q faltan
            sqlSp = "spCn_GeneraDatosCoa '" & CE(lrsProvision("Emp_cCodigo").Value) & "', '" & CE(lrsProvision("Pan_cAnio").Value) & "', '" & CE(lrsProvision("Per_cPeriodo").Value) _
                    & "', '" & CE(lrsProvision("Lib_cTipoLibro").Value) & "', '" & CE(lrsProvision("Ase_nVoucher").Value) & "', '" & CE(lrsProvision("Ase_cNummov").Value) & "', '" & CE(lrsProvision("Ten_cTipoEntidad").Value) _
                    & "', '" & CE(lrsProvision("Ent_cCodEntidad").Value) & "', '" & CE(lrsProvision("Asd_cTipoDoc").Value) & "', '" & CE(lrsProvision("Asd_cSerieDoc").Value) & "', '" & CE(lrsProvision("Asd_cNumDoc").Value) & "', '" & Me.tdbcMoneda.BoundText & "' "
            arrDatos = Array(sqlSp)

            Set clDatos = New clsMantoTablas
            Set rsArreglo = New ADODB.Recordset
            Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
            If Not rsArreglo Is Nothing Then
                If rsArreglo.State = adStateOpen Then
                    lArrDatos(i, 12) = NE(rsArreglo("fob").Value)           ' *** fob
                    lArrDatos(i, 13) = NE(rsArreglo("flete").Value)         ' *** flete
                End If
            End If

            Call CerrarRecordSet(rsArreglo)
            lrsProvision.MoveNext
        
            i = i + 1
        Loop
        Screen.MousePointer = vbDefault
        
        If Grabar = True Then
            Mensajes "Los datos se insertaron correctamente", vbInformation
         Else
            Mensajes "No se pudo insertar las importaciones.", vbInformation + vbOKOnly
         End If
          
        llenaGrilla
        
        Screen.MousePointer = vbDefault
        

        
    Else
        Mensajes "No se encontraron movimientos para el mes y libro seleccionado", vbInformation
        
    End If
    Call CerrarRecordSet(lrsProvision)

    Me.TDBGrid1.ReBind
    ' ***
    DoEvents
    Me.TDBGrid1.Refresh
    On Error Resume Next
    If i >= 0 Then TDBGrid1.Bookmark = 0
    
    DoEvents
    cmdTodos.Enabled = True
End Sub

Private Sub Form_Activate()
tdbcLibro.BoundText = lsLibroVen
If Sw = True Then Exit Sub
 If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
Sw = True
End Sub

Private Sub pCargaCfgLibro()
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    sqlver = "SELECT * From CNT_CONFIG_LIBROS WHERE Emp_cCodigo = '" & gsEmpresa & "'"
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
       lsLibroCom = CE(rsArreglo("Cfl_cCompras"))
       lsLibroVen = CE(rsArreglo("Cfl_cVentas"))
'       lsLibroHon = NuloText(rsArreglo("Cfl_cHonorarios"))
'       lsLibroDiario = NuloText(rsArreglo("Cfl_cDiario"))
'
'       If Len(Trim(NuloText(rsArreglo("Cfl_cCaja")))) > 0 Then
'          lsLibroCajIng = NuloText(rsArreglo("Cfl_cCaja"))
'          lsLibroCajEgr = NuloText(rsArreglo("Cfl_cCaja"))
'       Else
'          lsLibroCajIng = NuloText(rsArreglo("Cfl_cCajaIngresos"))
'          lsLibroCajEgr = NuloText(rsArreglo("Cfl_cCajaEgresos"))
'       End If
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Sub

Private Sub Form_Load()
    NUM_FILAS = 2000
    NUM_COLUMNAS = 30

    Dim sqlcombos As String
    Dim registros As Integer
    
    Sw = False
    pCargaCfgLibro
    
    Me.Top = (frmMDIConta.ScaleHeight - Me.Height) / 2
    Me.Left = (frmMDIConta.ScaleWidth - Me.Width) / 2
    
    Call LlenaComboMesAddItem(tdbcMes)

    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos

    
    If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
      
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                " WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' ORDER BY LIB_CDESCRIPCION "
    registros = LlenarComboAddItem(tdbcLibro, sqlcombos)
     
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        DesactivaBotones False
        TDBGrid1.Splits(0).Locked = True
    Else
        DesactivaBotones True
        TDBGrid1.Splits(0).Locked = False
    End If
    
    If registros > 0 Then
        DoEvents
        llenaGrilla
    
        tdbcLibro.Enabled = True
        tdbcLibro.Bookmark = 1
        tdbcLibro.ReBind
    Else
        Mensajes "No se crearon los libros contables en el sistema, ingreselos en mantenimiento de libros", vbOKOnly + vbInformation
        DesactivaBotones False
    End If
    

    Me.TDBGrid1.Columns(8).Visible = False
    Me.TDBGrid1.Columns(8).Width = 0
    Me.TDBGrid1.Columns(14).Visible = False
    Me.TDBGrid1.Columns(14).Width = 0
    Me.TDBGrid1.Columns(10).Visible = False
    Me.TDBGrid1.Columns(10).Width = 0
    Me.TDBGrid1.Columns(11).Visible = False
    Me.TDBGrid1.Columns(11).Width = 0
    
    lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS
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
    Dim sqlcombos As String
    Dim rsArreglo As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    Dim i As Integer
    Dim Col As Integer
    
    sqlcombos = "spCn_ConsultaAsientoCoa 'SEL_REGTIPOEXP',  '" & gsEmpresa & "',  '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '' "
'=========================================================================

    arrDatos = Array(sqlcombos)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo Is Nothing Then
        'Mensajes "Registro no existe. Seleccione un registro", vbInformation
        Screen.MousePointer = vbNormal
        Set rsArreglo = Nothing
        lArrDatos.Clear
        Set TDBGrid1.Array = lArrDatos
        TDBGrid1.Refresh
        
        Exit Sub
    End If
    
    lArrDatos.Clear
    
    If rsArreglo.RecordCount > NUM_FILAS Then
        NUM_FILAS = NUM_FILAS + 100
    End If
    
    lArrDatos.ReDim 0, NUM_FILAS, 0, NUM_COLUMNAS  ' filas
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
        lArrDatos(i, 8 + Col) = CE(rsArreglo(7))
        lArrDatos(i, 9 + Col) = CE(rsArreglo(8))
        lArrDatos(i, 10 + Col) = CE(rsArreglo(9).Value)
        lArrDatos(i, 11 + Col) = CE(rsArreglo(10).Value)
        lArrDatos(i, 12 + Col) = CE(rsArreglo(11).Value)
        lArrDatos(i, 13 + Col) = CE(rsArreglo(12).Value)
        lArrDatos(i, 14 + Col) = CE(rsArreglo(13).Value)
        lArrDatos(i, 15 + Col) = CE(rsArreglo(14).Value)
        lArrDatos(i, 16 + Col) = CE(rsArreglo(15).Value)
        lArrDatos(i, 17 + Col) = CE(rsArreglo(16).Value)
        lArrDatos(i, 18 + Col) = CE(rsArreglo(17).Value)
        lArrDatos(i, 19 + Col) = rsArreglo(18).Value
        lArrDatos(i, 20 + Col) = rsArreglo(19).Value
        lArrDatos(i, 21 + Col) = rsArreglo(20).Value
        lArrDatos(i, 22 + Col) = rsArreglo(21).Value
        lArrDatos(i, 23 + Col) = CE(rsArreglo(22).Value)
        lArrDatos(i, 24 + Col) = CE(rsArreglo(23).Value)
        lArrDatos(i, 25 + Col) = CE(rsArreglo(24).Value)
        lArrDatos(i, 26 + Col) = CE(rsArreglo(25).Value)
        lArrDatos(i, 27 + Col) = CE(rsArreglo(26).Value)
        lArrDatos(i, 28 + Col) = NuloNum(rsArreglo(27))
        lArrDatos(i, 29 + Col) = CE(rsArreglo(28).Value)
        lArrDatos(i, 30 + Col) = CE(rsArreglo(29).Value)
        
        
        
        rsArreglo.MoveNext
        
        i = i + 1
    Loop
    Set TDBGrid1.Array = lArrDatos
    TDBGrid1.Update
    DoEvents
    TDBGrid1.Refresh
    
    
    

'=========================================================================
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String)
   ' *** Dependiendo del control
    Dim psql$
    Select Case lControl
    Case "Provisiones"
        LlenaProvision
        'If TDBGrid1.Columns(0).Value = "" Then AñadeItem
        'Call SumarTotales
        Unload frmBusCoa
        pSetFocus Me.TDBGrid1
    End Select
End Sub

Private Sub LlenaProvision()
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim i As Integer
    Dim Fila As Integer
    i = 0
    Fila = CuentaFilas 'lArrDatos.Count(1) + 1
   
    With frmBusCoa.tdbgProvisiones
        lArrDatos(Fila, 0 + i) = (.Columns(7).Value) ' *** Tipo
        lArrDatos(Fila, 1 + i) = (.Columns(8).Value) ' *** Codigo
        lArrDatos(Fila, 2 + i) = (.Columns(9).Value) ' *** Ruc
        lArrDatos(Fila, 3 + i) = (.Columns(10).Value) ' *** Razon Social
        lArrDatos(Fila, 4 + i) = (.Columns(11).Value) ' *** Td
        lArrDatos(Fila, 5 + i) = (.Columns(12).Value) ' *** Serie
        lArrDatos(Fila, 6 + i) = (.Columns(13).Value) ' *** Numero
        lArrDatos(Fila, 7 + i) = (.Columns(17).Value) ' *** fecha
        lArrDatos(Fila, 8 + i) = ""                  ' *** Tipo Exportacion
        lArrDatos(Fila, 9 + i) = (.Columns(14).Value) ' *** Monto
        
        'If tdbcMoneda.BoundText = gsMonedaNac Then
            lArrDatos(Fila, 9 + i) = (.Columns(14).Value) ' *** Monto
        'Else
        '    lArrDatos(Fila, 9 + I) = (.Columns(16).Value) ' *** Monto
        'End If
        
        lArrDatos(Fila, 10 + i) = (.Columns(15).Value) ' *** Tc
        lArrDatos(Fila, 11 + i) = "" ' (.Columns(16).Value) ' *** Dolares
        
        lArrDatos(Fila, 17 + i) = (.Columns(19).Value) ' *** Glosa
        lArrDatos(Fila, 18 + i) = (.Columns(0).Value) ' *** Interno
        lArrDatos(Fila, 19 + i) = (.Columns(1).Value) ' *** Empresa
        lArrDatos(Fila, 20 + i) = (.Columns(2).Value) ' *** Año
        lArrDatos(Fila, 21 + i) = (.Columns(3).Value) ' *** Periodo
        lArrDatos(Fila, 22 + i) = (.Columns(4).Value) ' *** Libro
        lArrDatos(Fila, 23 + i) = (.Columns(5).Value) ' *** Voucher
        ' *** Aqui llamar a un sp q busq el valor de los datods q faltan
        sqlSp = "spCn_GeneraDatosCoa '" & (.Columns(1).Value) & "', '" & (.Columns(2).Value) & "', '" & (.Columns(3).Value) _
        & "', '" & (.Columns(4).Value) & "', '" & (.Columns(5).Value) & "', '" & (.Columns(0).Value) & "', '" & (.Columns(7).Value) _
        & "', '" & (.Columns(8).Value) & "', '" & (.Columns(11).Value) & "', '" & (.Columns(12).Value) & "', '" & (.Columns(13).Value) & "', '" & tdbcMoneda.BoundText & "' "
        arrDatos = Array(sqlSp)
        Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
        If rsArreglo.State = 0 Then Exit Sub
        lArrDatos(Fila, 12) = rsArreglo("fob").Value   ' *** fob
        lArrDatos(Fila, 13) = rsArreglo("flete").Value ' *** flete
        Call CerrarRecordSet(rsArreglo)
    End With
    
    DoEvents
    Grabar
    
    Set TDBGrid1.Array = lArrDatos
    TDBGrid1.Refresh

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        TDBGrid1.Width = Me.Width - 200
        TDBGrid1.Height = Me.Height - 1600
    End If
    
    Exit Sub
    
serror:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lArrDatos = Nothing
End Sub

Private Sub tdbcMes_ItemChange()

    llenaGrilla

End Sub

Private Sub TDBGrid1_BeforeRowColChange(Cancel As Integer)
    With TDBGrid1
        If (.Col = 7 Or .Col = 15 Or .Col = 16) Then
            If .Columns(.Col) <> "__/__/____" Then
                ' *** Si fecha no esta completa, completarla
                '.Columns(.col) = FormatoFecha(.Columns(.col))
                If VerificaFecha(.Columns(.Col)) = False Then
                    .RefreshRow
                    .SetFocus
                    Cancel = 1
                    .Columns(.Col) = "__/__/____"
                End If
            End If
        End If
    End With
End Sub

