VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcExportarPDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos PDB"
   ClientHeight    =   3945
   ClientLeft      =   1410
   ClientTop       =   1260
   ClientWidth     =   7530
   Icon            =   "frmPrcExportarCoa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   7530
   Begin VB.Frame Frame3 
      Height          =   3885
      Left            =   4320
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   2085
         _ExtentX        =   3678
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
         _PropDict       =   $"frmPrcExportarCoa.frx":1982
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
      Begin TrueOleDBList70.TDBCombo tdbcTipoPDB 
         Height          =   300
         Left            =   630
         TabIndex        =   11
         Top             =   765
         Width           =   2085
         _ExtentX        =   3678
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
         _PropDict       =   $"frmPrcExportarCoa.frx":1A09
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "TIPO DE DATOS A EXPORTAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   345
         TabIndex        =   10
         Top             =   315
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   9
         Top             =   1440
         Width           =   2205
      End
      Begin MSForms.CommandButton cmdExportar 
         Height          =   570
         Left            =   600
         TabIndex        =   4
         Top             =   2760
         Width           =   2115
         Caption         =   "   Exportar Datos"
         PicturePosition =   327683
         Size            =   "3731;1005"
         Picture         =   "frmPrcExportarCoa.frx":1A90
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4140
      Begin VB.TextBox tdbtDirectorio 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3360
         Width           =   3525
      End
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   840
         Width           =   3405
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DESTINO DE ARCHIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directorio seleccionado:"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   2130
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPrcExportarpdb"
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
    Dim respuesta As String
    
    ' *** Confirmando exportación
    respuesta = MsgBox("Desea exportar la información seleccionada", vbYesNo + vbQuestion, "Confirmar Exportar Datos")
    If respuesta = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    If CopiaArchivosCOA(CE(tdbtDirectorio)) = True Then
    
        ExportaEntidades
        TransfiereImportaciones
        TransfiereExportaciones
        TransfierePagos
        TransfiereNotas
    
        Mensajes "Los Datos se exportaron correctamente", vbInformation
    Else
        Mensajes "No se termino la exportacion correctamente", vbInformation
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Me.Top = (frmMDIConta.ScaleHeight - Me.Height) / 2
    Me.Left = (frmMDIConta.ScaleWidth - Me.Width) / 2
    
    Call LlenaComboMesAddItem(tdbcMes)
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    Dir1.Refresh
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdExportar.Enabled = False

    Else
        Me.cmdExportar.Enabled = True
        
    End If
    
End Sub

Private Sub Dir1_Change()
    tdbtDirectorio.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    tdbtDirectorio.Text = Dir1.Path
End Sub

Private Sub ExportaEntidades()
    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    
    sqlEnt = "Select Ent_nRuc, Ent_cCodEntidad, Ent_cPersona From dbo.CNM_ENTIDAD "
    sqlEnt = sqlEnt + " Where Emp_cCodigo = '" & gsEmpresa & "' and Ten_cTipoEntidad = 'P' and Ent_cTipoDoc = '04' "
    Set rsDatos = ConsultarDatosRs(sqlEnt)
    ruta = Trim(tdbtDirectorio) '+ "\X61COAPR.DBF"
    rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    cnExp.ConnectionString = rutaExp
    'cnExp.ConnectionTimeout = 30
    cnExp.Open
    i = 0
    If Not rsDatos Is Nothing Then
        ' *** Ingresar registros a TABLA DBF
        Do While Not rsDatos.EOF
            i = i + 1
            sqlEnt = " Insert into X61COAPR.DBF (X61NROREG, X61NRORUC, X61CINTER, X61NOMBRE, X61ESTADO) "
            sqlEnt = sqlEnt + " Values ( " & i & ", '" & rsDatos("Ent_nRuc") & "', '" & rsDatos("Ent_cCodEntidad") & "', '" & rsDatos("Ent_cPersona") & "', '1') "
            cnExp.Execute sqlEnt
            rsDatos.MoveNext
        Loop
        ' ***
    End If
    Call CerrarRecordSet(rsDatos)
    cnExp.Close
    Set cnExp = Nothing
End Sub

Private Sub TransfiereImportaciones()
    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim periodo As String
    Dim fpago As String
    Dim tipoCompra As String
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    
    sqlEnt = "Select * From dbo.CND_ASIENTO_COA "
    sqlEnt = sqlEnt + " Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
    sqlEnt = sqlEnt + " AND Per_cPeriodo = '" & tdbcMes.BoundText & "' AND Dco_cTipo = 'I'"
    Set rsDatos = ConsultarDatosRs(sqlEnt)
    ruta = Trim(tdbtDirectorio) '+ "\X78CPCOA.DBF"
    rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    cnExp.ConnectionString = rutaExp
    'cnExp.ConnectionTimeout = 30
    cnExp.Open
    i = 0
    If Not rsDatos Is Nothing Then
        ' *** Ingresar registros a TABLA DBF
        Do While Not rsDatos.EOF
            i = i + 1
            periodo = Format(rsDatos("Pan_cAnio"), "0000") + Format(rsDatos("Per_cPeriodo"), "00")
            
            fpago = ""
            If CE(rsDatos("Dco_cFechaPago")) <> "" Then
                fpago = Format(Year(rsDatos("Dco_cFechaPago")), "0000") + Format(Month(rsDatos("Dco_cFechaPago")), "00") + Format(Day(rsDatos("Dco_cFechaPago")), "00")
            End If
            
            Select Case rsDatos("Dco_cTipoIGV")
                Case "A"
                    tipoCompra = "001"
                Case "B"
                    tipoCompra = "002"
                Case "C"
                    tipoCompra = "003"
                Case Else
                    tipoCompra = "004"
            End Select
            sqlEnt = " Insert into X78CPCOA.DBF "
            sqlEnt = sqlEnt + " (X78NROREG, X78PERIODO, X78F_PAGO, X78TIPO, X78NUMPOL, X78NUMFRAC, X78NUMCUO, X78M_IMPOR, X78ADVALOR, X78ISC, X78IGV, X78CODIGO, X78NORDEN, X78ESTADO) "
            sqlEnt = sqlEnt + " Values ( " & i & ", '" & periodo & "', '" & fpago & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', '" & RTrim(rsDatos("Dco_cNumDoc")) & "', '', "
            sqlEnt = sqlEnt + " 0, " & rsDatos("Dco_nMonto") & ", " & rsDatos("Dco_nMontoAdvalorem") & ", " & rsDatos("Dco_nMontoISC") & ",  "
            sqlEnt = sqlEnt + " " & rsDatos("Dco_nMontoIGV") & ", '" & tipoCompra & "', '" & i & "', '1') "
            cnExp.Execute sqlEnt
            rsDatos.MoveNext
        Loop
        ' ***
    End If
    Call CerrarRecordSet(rsDatos)
    cnExp.Close
    Set cnExp = Nothing
End Sub

Private Sub TransfiereExportaciones()
    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim periodo As String
    Dim fechaDoc As String
    Dim fechaDue As String
    Dim fechaEmb As String
    Dim Moneda As String
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    
    sqlEnt = "Select * From CND_ASIENTO_COA "
    sqlEnt = sqlEnt + " Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
    sqlEnt = sqlEnt + " AND Per_cPeriodo = '" & tdbcMes.BoundText & "' AND Dco_cTipo = 'E'"
    Set rsDatos = ConsultarDatosRs(sqlEnt)
    ruta = Trim(tdbtDirectorio) '+ "\X78CPCOA.DBF"
    rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    cnExp.ConnectionString = rutaExp
    'cnExp.ConnectionTimeout = 30
    cnExp.Open
    i = 0
    If Not rsDatos Is Nothing Then
        ' *** Ingresar registros a TABLA DBF
        Do While Not rsDatos.EOF
            i = i + 1
            periodo = Format(rsDatos("Pan_cAnio"), "0000") + Format(rsDatos("Per_cPeriodo"), "00")
            fechaDoc = Format(Year(rsDatos("Dco_dFecDoc")), "0000") + Format(Month(rsDatos("Dco_dFecDoc")), "00") + Format(Day(rsDatos("Dco_dFecDoc")), "00")
            If CE(rsDatos("Dco_dFechaNumera")) <> "" Then
                fechaDue = Format(Year(rsDatos("Dco_dFechaNumera")), "0000") + Format(Month(rsDatos("Dco_dFechaNumera")), "00") + Format(Day(rsDatos("Dco_dFechaNumera")), "00")
            Else
                fechaDue = ""
            End If
            If CE(rsDatos("Dco_dFechaEmbarque")) <> "" Then
                fechaEmb = Format(Year(rsDatos("Dco_dFechaEmbarque")), "0000") + Format(Month(rsDatos("Dco_dFechaEmbarque")), "00") + Format(Day(rsDatos("Dco_dFechaEmbarque")), "00")
            Else
                fechaEmb = ""
            End If
            If monedaNacional(rsDatos("Dco_cMoneda")) = True Then
                Moneda = "1"
            Else
                Moneda = "0"
            End If
            sqlEnt = " Insert into X88CPCOA.DBF "
            sqlEnt = sqlEnt + " (X88NROREG, X88PERIODO, X88TIPEXP, X88FECHA, X88TIPO, X88SERIE, X88NUMERO, X88NUMDUE, X88F_DUE, X88F_EMB, X88FOB, "
            sqlEnt = sqlEnt + " X88TIPCAM, X88FLAGMON, X88NORDEN, X88TIPORI, X88SERORI, X88NUMORI, X88FECORI, X88ESTADO) "
            sqlEnt = sqlEnt + " Values ( " & i & ", '" & periodo & "', '" & rsDatos("Dco_cTipoBS") & "', '" & fechaDoc & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', "
            sqlEnt = sqlEnt + " '" & rsDatos("Dco_cSerieDoc") & "', '" & rsDatos("Dco_cNumDoc") & "', '" & rsDatos("Dco_nNumDue") & "', '" & fechaDue & "', "
            sqlEnt = sqlEnt + " '" & fechaEmb & "', " & rsDatos("Dco_nValorFOB") & ", " & rsDatos("Asd_nTipoCambio") & ", '" & Moneda & "', "
            sqlEnt = sqlEnt + " '" & i & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', '" & rsDatos("Dco_cSerieDoc") & "', '" & rsDatos("Dco_cNumDoc") & "', '" & fechaDoc & "', '1') "
            cnExp.Execute sqlEnt
            rsDatos.MoveNext
        Loop
        ' ***
    End If
    Call CerrarRecordSet(rsDatos)
    cnExp.Close
    Set cnExp = Nothing
End Sub

Private Sub TransfierePagos()
    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim periodo As String
    Dim fpago As String
    Dim tipoCompra As String
    Dim BASE As Double
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    
    sqlEnt = "Select * From dbo.CND_ASIENTO_COA "
    sqlEnt = sqlEnt + " Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
    sqlEnt = sqlEnt + " AND Per_cPeriodo = '" & tdbcMes.BoundText & "' AND Dco_cTipo = 'I'"
    sqlEnt = sqlEnt + " AND Dco_cTipoDoc NOT IN (SELECT Tab_cCodigo FROM TABLA "
    sqlEnt = sqlEnt + " WHERE Emp_cCodigo = '" & gsEmpresa & "' AND tab_ctabla in ('033','034')) "
    
    Set rsDatos = ConsultarDatosRs(sqlEnt)
    ruta = Trim(tdbtDirectorio) '+ "\X78CPCOA.DBF"
    rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    cnExp.ConnectionString = rutaExp
    'cnExp.ConnectionTimeout = 30
    cnExp.Open
    i = 0
    If Not rsDatos Is Nothing Then
        ' *** Ingresar registros a TABLA DBF
        Do While Not rsDatos.EOF
            i = i + 1
            periodo = Format(rsDatos("Pan_cAnio"), "0000") + Format(rsDatos("Per_cPeriodo"), "00")
            
            fpago = ""
            If CE(rsDatos("Dco_cFechaPago")) <> "" Then
                fpago = Format(Year(rsDatos("Dco_dFecDoc")), "0000") + Format(Month(rsDatos("Dco_dFecDoc")), "00") + Format(Day(rsDatos("Dco_dFecDoc")), "00")
            End If
            
            BASE = rsDatos("Dco_nMonto") - (rsDatos("Dco_nMontoIGV") + rsDatos("Dco_nMontoAdvalorem") + rsDatos("Dco_nMontoISC") + rsDatos("Dco_nMontoIPM"))
            Select Case rsDatos("Dco_cTipoIGV")
                Case "A"
                    tipoCompra = "001"
                Case "B"
                    tipoCompra = "002"
                Case "C"
                    tipoCompra = "003"
                Case Else
                    tipoCompra = "004"
            End Select
            sqlEnt = " Insert into X58CPCOA.DBF "
            sqlEnt = sqlEnt + " (X58NROREG, X58IDENT, X58PERIODO, X58FECHA, X58TIPO, X58SERIE, X58NUMERO, X58BASE, X58IGV, X58CODIGO, X58FLAGMON, X58NORDEN, X58ESTADO) "
            sqlEnt = sqlEnt + " Values ( " & i & ", '" & Trim(rsDatos("Dco_cNumRuc")) & "', '" & periodo & "', '" & fpago & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', '" & RTrim(rsDatos("Dco_cSerieDoc")) & "','" & RTrim(rsDatos("Dco_cNumDoc")) & "', "
            sqlEnt = sqlEnt + " " & BASE & ", " & rsDatos("Dco_nMontoIGV") & ", '" & tipoCompra & "', '1', '" & i & "', '1') "
            cnExp.Execute sqlEnt
            rsDatos.MoveNext
        Loop
        ' ***
    End If
    Call CerrarRecordSet(rsDatos)
    cnExp.Close
    Set cnExp = Nothing
End Sub

Private Sub TransfiereNotas()
    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim periodo As String
    Dim fpago As String
    Dim tipoCompra As String
    Dim BASE As Double
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    
    sqlEnt = "Select * From dbo.CND_ASIENTO_COA "
    sqlEnt = sqlEnt + " Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
    sqlEnt = sqlEnt + " AND Per_cPeriodo = '" & tdbcMes.BoundText & "' AND Dco_cTipo = 'I'"
    sqlEnt = sqlEnt + " AND ( Dco_cTipoDoc = '07' or Dco_cTipoDoc = '08') "

    
    Set rsDatos = ConsultarDatosRs(sqlEnt)
    ruta = Trim(tdbtDirectorio) '+ "\X78CPCOA.DBF"
    rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    cnExp.ConnectionString = rutaExp
   ' cnExp.ConnectionTimeout = 30
    cnExp.Open
    i = 0
    If Not rsDatos Is Nothing Then
        ' *** Ingresar registros a TABLA DBF
        Do While Not rsDatos.EOF
            i = i + 1
            periodo = Format(rsDatos("Pan_cAnio"), "0000") + Format(rsDatos("Per_cPeriodo"), "00")
            
            fpago = ""
            If CE(rsDatos("Dco_cFechaPago")) <> "" Then
                fpago = Format(Year(rsDatos("Dco_dFecDoc")), "0000") + Format(Month(rsDatos("Dco_dFecDoc")), "00") + Format(Day(rsDatos("Dco_dFecDoc")), "00")
            End If
            BASE = rsDatos("Dco_nMonto") - (rsDatos("Dco_nMontoIGV") + rsDatos("Dco_nMontoAdvalorem") + rsDatos("Dco_nMontoISC") + rsDatos("Dco_nMontoIPM"))
            Select Case rsDatos("Dco_cTipoIGV")
                Case "A"
                    tipoCompra = "001"
                Case "B"
                    tipoCompra = "002"
                Case "C"
                    tipoCompra = "003"
                Case Else
                    tipoCompra = "004"
            End Select
            sqlEnt = " Insert into X68CPCOA.DBF "
            sqlEnt = sqlEnt + " (X68NROREG, X68IDENT, X68PERIODO, X68FECHA, X68TIPO, X68SERIE, X68NUMERO, X68BASE, X68IGV, X68CODIGO, X68FLAGMON, X68NORDEN, "
            sqlEnt = sqlEnt + " X68TIPORI, X68SERORI, X68NUMORI, X68FECORI, X68ESTADO) "
            sqlEnt = sqlEnt + " Values ( " & i & ", '" & Trim(rsDatos("Dco_cNumRuc")) & "', '" & periodo & "', '" & fpago & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', '" & RTrim(rsDatos("Dco_cSerieDoc")) & "', '" & RTrim(rsDatos("Dco_cNumDoc")) & "', "
            sqlEnt = sqlEnt + " " & BASE & ", " & rsDatos("Dco_nMontoIGV") & ", '" & tipoCompra & "', '1', '" & i & "', '" & Trim(rsDatos("Dco_cTipoDoc")) & "', '" & RTrim(rsDatos("Dco_cSerieDoc")) & "', '" & RTrim(rsDatos("Dco_cNumDoc")) & "', '" & fpago & "', '1') "
            cnExp.Execute sqlEnt
            rsDatos.MoveNext
        Loop
        ' ***
    End If
    Call CerrarRecordSet(rsDatos)
    cnExp.Close
    Set cnExp = Nothing
End Sub
