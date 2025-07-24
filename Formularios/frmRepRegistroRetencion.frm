VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepRegistroRetencion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Retenciones y Percepciones"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   Icon            =   "frmRepRegistroRetencion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   6780
   Begin VB.Frame fraTodo 
      Height          =   3690
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   6645
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   435
         Left            =   4320
         TabIndex        =   16
         Top             =   3000
         Width           =   1665
      End
      Begin VB.Frame Frame2 
         Height          =   1320
         Left            =   345
         TabIndex        =   5
         Top             =   240
         Width           =   5940
         Begin VB.OptionButton optPercepcion 
            Caption         =   "REGISTRO DE PERCEPCION"
            Height          =   195
            Left            =   2535
            TabIndex        =   1
            Top             =   675
            Width           =   2535
         End
         Begin VB.OptionButton optRetencion 
            Caption         =   "REGISTRO DE RETENCION"
            Height          =   195
            Left            =   2535
            TabIndex        =   0
            Top             =   315
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Seleccione el Registro que desea Imprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1710
         End
      End
      Begin TrueOleDBList70.TDBCombo tdbcMoneda 
         Height          =   300
         Left            =   3098
         TabIndex        =   3
         Tag             =   "_"
         Top             =   2355
         Width           =   1980
         _ExtentX        =   3493
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
         _PropDict       =   $"frmRepRegistroRetencion.frx":0ECA
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
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   3143
         TabIndex        =   2
         Top             =   1920
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
         _PropDict       =   $"frmRepRegistroRetencion.frx":0F51
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
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   2520
         TabIndex        =   14
         Top             =   3000
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepRegistroRetencion.frx":0FD8
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   720
         TabIndex        =   13
         Top             =   3000
         Width           =   1665
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepRegistroRetencion.frx":1572
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MONEDA"
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
         Left            =   1703
         TabIndex        =   8
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MES"
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
         Left            =   1703
         TabIndex        =   7
         Top             =   1935
         Width           =   375
      End
   End
   Begin TDBDate6Ctl.TDBDate dtpDesde 
      Height          =   300
      Left            =   3825
      TabIndex        =   9
      Tag             =   "enabled"
      Top             =   4350
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   529
      Calendar        =   "frmRepRegistroRetencion.frx":1B0C
      Caption         =   "frmRepRegistroRetencion.frx":1C0E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepRegistroRetencion.frx":1C72
      Keys            =   "frmRepRegistroRetencion.frx":1C90
      Spin            =   "frmRepRegistroRetencion.frx":1CFC
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
      Left            =   3825
      TabIndex        =   10
      Tag             =   "enabled"
      Top             =   4755
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   529
      Calendar        =   "frmRepRegistroRetencion.frx":1D24
      Caption         =   "frmRepRegistroRetencion.frx":1E26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepRegistroRetencion.frx":1E8A
      Keys            =   "frmRepRegistroRetencion.frx":1EA8
      Spin            =   "frmRepRegistroRetencion.frx":1F14
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
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE DE COMPROBACION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   4365
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
      Left            =   2625
      TabIndex        =   12
      Top             =   4380
      Visible         =   0   'False
      Width           =   630
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
      Left            =   2625
      TabIndex        =   11
      Top             =   4785
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmRepRegistroRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdExportar_Click()
    
    Dim strSql, strSqlAux1, strSqlAux2, strSqlAux3 As String
    Dim strTipo As String
    Dim aDatos() As Variant
    Dim objDatos As clsMantoTablas
    Dim ObjFuncion As New ClsFuncionesExecute
    Dim rsExportar As ADODB.Recordset
    Dim rsExportarAux1 As ADODB.Recordset
    Dim rsExportarAux2 As ADODB.Recordset
    Dim rsExportarAux3 As ADODB.Recordset
        
    Screen.MousePointer = vbHourglass
    
    Dim Mes As String
    Mes = tdbcMes.BoundText
    If Mes = "00" Then Mes = "01"
    If Mes > "13" Then Mes = "12"
    
    ' *** Aqui hallar la fecha de inicio y la fecha de fin
    dtpDesde.Value = "01/" + Mes + "/" + gsAnio
    dtpHasta.Value = UltimoDiaMes(Mes, gsAnio)
    
    If Me.optRetencion.Value = True Then
'        strSql = "spCn_ExportarRetenciones 'RETENCIONES', '', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '', '', '" & CE(dtpDesde.Value) & "', '" & CE(dtpHasta.Value) & "', '', '" & Me.tdbcMoneda.BoundText & "'"
        strSql = "USP_Exportacion_Retencion_PDT '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcMoneda.BoundText & "'"
        strSqlAux3 = "USP_Exportar_Retencion_PDT_621 '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcMoneda.BoundText & "'"
    ElseIf Me.optPercepcion.Value = True Then
'        strSql = "spCn_ExportarPercepciones 'PERCEPCIONES', '', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "' ,'', '', '" & CE(dtpDesde.Value) & "', '" & CE(dtpHasta.Value) & "', '', '" & Me.tdbcMoneda.BoundText & "'"
        strSql = "USP_Exportar_Percepcion_PDT '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcMoneda.BoundText & "'"
        strSqlAux1 = "USP_Exportar_Percepcion_Ventas '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcMoneda.BoundText & "'"
        strSqlAux2 = "USP_Exportar_Percepcion_PDT_621 '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcMes.BoundText & "', '" & Me.tdbcMoneda.BoundText & "'"
    End If
    
    Set objDatos = New clsMantoTablas
    Set rsExportar = New ADODB.Recordset
    Set rsExportarAux1 = New ADODB.Recordset
    Set rsExportarAux2 = New ADODB.Recordset
    Set rsExportarAux3 = New ADODB.Recordset
    
    Set rsExportar = ObjFuncion.fRetornaRS(strSql)
    
    If Me.optPercepcion.Value = True Then
        Set rsExportarAux1 = ObjFuncion.fRetornaRS(strSqlAux1)
        Set rsExportarAux2 = ObjFuncion.fRetornaRS(strSqlAux2)
    End If
    
    If Me.optRetencion.Value = True Then
        Set rsExportarAux3 = ObjFuncion.fRetornaRS(strSqlAux3)
    End If
  
    Dim strRuta As String
    Dim strRutaAux1, strRutaAux2, strRutaAux3 As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strRuta = App.Path + "\PDT\"
    strRutaAux1 = App.Path + "\PDT\"
    
    If fso.FolderExists(strRuta) = False Then
        fso.CreateFolder (strRuta)
        strRuta = strRuta & IIf(Me.optPercepcion.Value = True, gsEmpresa & "-PERCEPCION\", gsEmpresa & "-RETENCION\")
        If fso.FolderExists(strRuta) = False Then
            fso.CreateFolder (strRuta)
        End If
    Else
        strRuta = strRuta & IIf(Me.optPercepcion.Value = True, gsEmpresa & "-PERCEPCION\", gsEmpresa & "-RETENCION\")
        If fso.FolderExists(strRuta) = False Then
            fso.CreateFolder (strRuta)
        End If
    End If
    
    If Me.optPercepcion.Value = True Then
        strRutaAux1 = strRuta
        strRutaAux2 = strRuta
    End If
    
    If Me.optRetencion.Value = True Then
        strRutaAux3 = strRuta
    End If
    
    If Me.optRetencion.Value = True Then
        strRuta = strRuta & "0626" & gsRUC & gsAnio & Me.tdbcMes.BoundText & ".txt"
        strRutaAux3 = strRutaAux3 & "0621" & gsRUC & gsAnio & Me.tdbcMes.BoundText & "R.txt"
    ElseIf Me.optPercepcion.Value = True Then
        strRuta = strRuta & "0633" & gsRUC & gsAnio & Me.tdbcMes.BoundText & ".txt"
        strRutaAux1 = strRutaAux1 & "0697" & gsRUC & gsAnio & Me.tdbcMes.BoundText & ".txt"
        strRutaAux2 = strRutaAux2 & "0621" & gsRUC & gsAnio & Me.tdbcMes.BoundText & "P.txt"
    End If
    
    Dim intIndex As Integer
    
    If rsExportar.RecordCount > 0 Then
        Open strRuta For Output Shared As #1
        Close #1
        
        Open strRuta For Append As #1
        
        For intIndex = 0 To rsExportar.RecordCount - 1
            Print #1, rsExportar.Fields("Registro").Value
            rsExportar.MoveNext
        Next intIndex
        Close #1
    End If
    
    If Me.optPercepcion.Value = True Then
        If rsExportarAux1.RecordCount > 0 Then
            Open strRutaAux1 For Output Shared As #1
            Close #1
            
            Open strRutaAux1 For Append As #1
            
            For intIndex = 0 To rsExportarAux1.RecordCount - 1
                Print #1, rsExportarAux1.Fields("Registro").Value
                rsExportarAux1.MoveNext
            Next intIndex
            Close #1
        End If
        
        If rsExportarAux2.RecordCount > 0 Then
            Open strRutaAux2 For Output Shared As #1
            Close #1
            
            Open strRutaAux2 For Append As #1
            
            For intIndex = 0 To rsExportarAux2.RecordCount - 1
                Print #1, rsExportarAux2.Fields("Registro").Value
                rsExportarAux2.MoveNext
            Next intIndex
            Close #1
        End If
        
    End If
    
    If Me.optRetencion.Value = True Then
        If rsExportarAux3.RecordCount > 0 Then
            Open strRutaAux3 For Output Shared As #1
            Close #1
            
            Open strRutaAux3 For Append As #1
            
            For intIndex = 0 To rsExportarAux3.RecordCount - 1
                Print #1, rsExportarAux3.Fields("Registro").Value
                rsExportarAux3.MoveNext
            Next intIndex
            Close #1
        End If
    End If
        
    If Me.optPercepcion.Value = True Then
        If rsExportar.RecordCount = 0 And rsExportarAux1.RecordCount = 0 And rsExportarAux2.RecordCount = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "No se encontro informacion para la Exportarcion", vbExclamation, "Sistema"
            Exit Sub
        End If
    End If
    
    If Me.optRetencion.Value = True Then
        If rsExportar.RecordCount = 0 And rsExportarAux3.RecordCount = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "No se encontro informacion para la Exportarcion", vbExclamation, "Sistema"
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbDefault
    MsgBox "La Exportacion se realizo satisfactoriamente en : " & strRuta, vbInformation, "Sistema"
    
End Sub

Private Sub cmdImprimir_Click()
    ' *** Abrir el reporte y enviar los parametros
    Dim matriz_fecha(13) As Variant
    Dim cadFecha As String
    Dim fecHasta As Date
    
    Dim Mes As String
    Mes = tdbcMes.BoundText
    
    If Mes = "00" Then Mes = "01"
    If Mes > "13" Then Mes = "12"
    
    
    Screen.MousePointer = vbHourglass

    ' *** Aqui hallar la fecha de inicio y la fecha de fin
    dtpDesde.Value = "01/" + Mes + "/" + gsAnio
    dtpHasta.Value = UltimoDiaMes(Mes, gsAnio)
    ' ***
    
    matriz_fecha(0) = "@Param_cCadena;" & "PERIODO " & tdbcMes.Text & " " & gsAnio & " EN " & tdbcMoneda.Text & ";True"
    
    matriz_fecha(2) = "@Ase_cNummov;;True"
    matriz_fecha(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(4) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz_fecha(5) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
    
    matriz_fecha(6) = "@Lib_cTipoLibro;;True"
    matriz_fecha(7) = "@Ase_nVoucher;;True"
    matriz_fecha(8) = "@desde;" & CE(dtpDesde.Value) & ";True"
    matriz_fecha(9) = "@hasta;" & CE(dtpHasta.Value) & ";True"
    matriz_fecha(10) = "@Pla_cCuentaContable;;True"
    matriz_fecha(11) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
    matriz_fecha(12) = "NombreLargo;" & gsEmpresaNom & ";True"
    matriz_fecha(13) = "@RUC;" & "RUC : " & gsRUC & ";True"

    Dim formulas(0) As Variant
    cmdImprimir.Enabled = False
    If Me.optRetencion.Value = True Then
        matriz_fecha(1) = "@Tipo;" & "RETENCIONES;True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptRegistroRetencion.rpt", crptToWindow, "Registro de Regimen de Retención", "", matriz_fecha(), formulas()
        
        
    ElseIf Me.optPercepcion.Value = True Then
        matriz_fecha(1) = "@Tipo;" & "PERCEPCIONES;True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptRegistroPercepcion.rpt", crptToWindow, "Registro de Regimen de Percepción", "", matriz_fecha(), formulas()
        
    End If
    
    
    Screen.MousePointer = vbNormal
    cmdImprimir.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    
    dtpHasta = FechaServidor
    dtpDesde = dtpHasta
    Call LlenaCombos
    
    
    tdbcMoneda.BoundText = gsMonedaNac
    
    If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
    
'    If gintRetencion = 1 Or gintPercepcion = 1 Then
'        Me.cmdExportar.Enabled = True
'    Else
'        Me.cmdExportar.Enabled = False
'    End If
    
    tdbcMes.ReBind
    tdbcMoneda.ReBind
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Call LlenaComboMesAddItem(tdbcMes)
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
End Sub

Private Sub BuscarMonedaNacional()
    Dim i As Integer
    For i = 0 To tdbcMoneda.ListCount - 1
        tdbcMoneda.Row = i
        If tdbcMoneda.Columns(2).Value = "1" Then
            tdbcMoneda.Bookmark = i
            Exit Sub
        End If
    Next
    tdbcMoneda.Bookmark = tdbcMoneda.Bookmark
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fraTodo, Me)
        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub


Private Sub optCompras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMes
End If
End Sub



Private Sub optHonorarios_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If tdbcMes.Enabled = False Then
        pSetFocus tdbcMoneda
    Else
        pSetFocus tdbcMes
    End If
End If
End Sub


Private Sub optVentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMes
End If
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMoneda
End If
End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub


