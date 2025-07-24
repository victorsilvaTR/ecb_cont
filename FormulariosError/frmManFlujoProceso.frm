VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManFlujoProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Proceso del Flujo de Efectivo - Hoja de Trabajo"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   Icon            =   "frmManFlujoProceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   11325
   Begin VB.Frame fraTodo 
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11265
      Begin TrueOleDBGrid70.TDBDropDown tdbdActividad 
         Height          =   1425
         Left            =   5670
         TabIndex        =   16
         Top             =   2565
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   2514
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFF9D7&,.bold=0,.fontsize=825"
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
         Height          =   1185
         Left            =   5490
         TabIndex        =   6
         Top             =   0
         Width           =   5640
         Begin TrueOleDBList70.TDBCombo tdbcMesImportar 
            Height          =   300
            Left            =   1710
            TabIndex        =   7
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
            _PropDict       =   $"frmManFlujoProceso.frx":0ECA
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
            Left            =   180
            TabIndex        =   9
            Top             =   315
            Width           =   660
         End
         Begin MSForms.CommandButton cmdImportar 
            Height          =   360
            Left            =   3780
            TabIndex        =   8
            ToolTipText     =   " Importar los datos del periodo seleccionado "
            Top             =   675
            Width           =   1305
            Caption         =   " Importar"
            PicturePosition =   327683
            Size            =   "2302;635"
            Picture         =   "frmManFlujoProceso.frx":0F51
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Periodo del configuración del proceso "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5010
         Begin VB.ComboBox cboMetodo 
            Height          =   315
            ItemData        =   "frmManFlujoProceso.frx":14EB
            Left            =   1170
            List            =   "frmManFlujoProceso.frx":14F5
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   675
            Width           =   3390
         End
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   1170
            TabIndex        =   3
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
            _PropDict       =   $"frmManFlujoProceso.frx":150D
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
            Left            =   180
            TabIndex        =   17
            Top             =   720
            Width           =   630
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
            Left            =   225
            TabIndex        =   4
            Top             =   315
            Width           =   660
         End
      End
      Begin TrueOleDBGrid70.TDBDropDown tdbdTipo 
         Height          =   1875
         Left            =   7470
         TabIndex        =   1
         Top             =   2565
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   3307
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
      Begin TrueOleDBGrid70.TDBGrid tdbgFlujo 
         Height          =   4065
         Left            =   0
         TabIndex        =   5
         Top             =   1755
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   7170
         _LayoutType     =   4
         _RowHeight      =   17
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cuenta"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripción"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "tdbdPlantilla"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "CodActividad"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Tipo Actividad"
         Columns(4).DataField=   ""
         Columns(4).DropDown=   "tdbdActividad"
         Columns(4).DropDown.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CodTipoDebe"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Ajuste Debe"
         Columns(6).DataField=   ""
         Columns(6).DropDown=   "tdbdTipo"
         Columns(6).DropDown.vt=   8
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "FormulaDebeCod"
         Columns(7).DataField=   ""
         Columns(7).DropDown=   "tdbdColumna"
         Columns(7).DropDown.vt=   8
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Formula o  Valor - Ajuste Debe"
         Columns(8).DataField=   ""
         Columns(8).DropDown=   "tdbdColumna"
         Columns(8).DropDown.vt=   8
         Columns(8).ButtonPicture.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(8).ButtonPicture(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(8).ButtonPicture(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(2)=   "7+vx7+vx7+vx7+vx7+trrYQhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(3)=   "7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU"
         Columns(8).ButtonPicture(4)=   "3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx"
         Columns(8).ButtonPicture(5)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(6)=   "7+vx7+vx7+vx7+trrYQhhCkhhCkhhCkhhCkhhCmU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx"
         Columns(8).ButtonPicture(7)=   "7+tjpWM5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVIhhCnx7+vx7+tjpWOU3oyU"
         Columns(8).ButtonPicture(8)=   "3oyU3oyU3oyU3oyU3ow5tVKU3oyU3oyU3oyU3oyU3owhhCnx7+vx7+trrYRjpWNjpWNjpWNjpWNj"
         Columns(8).ButtonPicture(9)=   "pWOU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIh"
         Columns(8).ButtonPicture(10)=   "hCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx"
         Columns(8).ButtonPicture(11)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(12)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(13)=   "7+vx7+vx7+vx7+trrYRjpWNjpWNrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(8).ButtonPicture(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+s="
         Columns(8).ButtonPicture.vt=   9
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "CodTipoHaber"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Ajuste Haber"
         Columns(10).DataField=   ""
         Columns(10).DropDown=   "tdbdTipo"
         Columns(10).DropDown.vt=   8
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "FormulaHaberCod"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Formula o  Valor - Ajuste Haber"
         Columns(12).DataField=   ""
         Columns(12).ButtonPicture.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(12).ButtonPicture(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(12).ButtonPicture(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(2)=   "7+vx7+vx7+vx7+vx7+trrYQhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(3)=   "7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU"
         Columns(12).ButtonPicture(4)=   "3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx"
         Columns(12).ButtonPicture(5)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(6)=   "7+vx7+vx7+vx7+trrYQhhCkhhCkhhCkhhCkhhCmU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx"
         Columns(12).ButtonPicture(7)=   "7+tjpWM5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVIhhCnx7+vx7+tjpWOU3oyU"
         Columns(12).ButtonPicture(8)=   "3oyU3oyU3oyU3oyU3ow5tVKU3oyU3oyU3oyU3oyU3owhhCnx7+vx7+trrYRjpWNjpWNjpWNjpWNj"
         Columns(12).ButtonPicture(9)=   "pWOU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIh"
         Columns(12).ButtonPicture(10)=   "hCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx"
         Columns(12).ButtonPicture(11)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(12)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+tjpWOU3ow5tVIhhCnx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(13)=   "7+vx7+vx7+vx7+trrYRjpWNjpWNrrYTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
         Columns(12).ButtonPicture(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+s="
         Columns(12).ButtonPicture.vt=   9
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).SizeMode=   2
         Splits(0).Size  =   3
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=139793"
         Splits(0)._ColumnProps(6)=   "Column(0).FetchStyle=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(0).Merge=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1535"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1455"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=139776"
         Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=4366"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4286"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=131588"
         Splits(0)._ColumnProps(19)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=139780"
         Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(28)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(30)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=139780"
         Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(36)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(38)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=139780"
         Splits(0)._ColumnProps(43)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(44)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(45)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(46)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=131588"
         Splits(0)._ColumnProps(51)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(52)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(53)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(54)=   "Column(6).AutoDropDown=1"
         Splits(0)._ColumnProps(55)=   "Column(6).DropDownList=1"
         Splits(0)._ColumnProps(56)=   "Column(7).Width=1244"
         Splits(0)._ColumnProps(57)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(7)._WidthInPix=1164"
         Splits(0)._ColumnProps(59)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(60)=   "Column(7)._ColStyle=131588"
         Splits(0)._ColumnProps(61)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(62)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(63)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(64)=   "Column(8).Width=7646"
         Splits(0)._ColumnProps(65)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(8)._WidthInPix=7567"
         Splits(0)._ColumnProps(67)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(68)=   "Column(8)._ColStyle=131588"
         Splits(0)._ColumnProps(69)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(70)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(71)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(72)=   "Column(8).AutoDropDown=1"
         Splits(0)._ColumnProps(73)=   "Column(8).DropDownList=1"
         Splits(0)._ColumnProps(74)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(75)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(77)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(78)=   "Column(9)._ColStyle=139780"
         Splits(0)._ColumnProps(79)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(80)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(81)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(82)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(83)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(84)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(85)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(86)=   "Column(10)._ColStyle=139780"
         Splits(0)._ColumnProps(87)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(88)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(89)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(90)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(91)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(92)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(93)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(94)=   "Column(11)._ColStyle=139780"
         Splits(0)._ColumnProps(95)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(96)=   "Column(11).AllowFocus=0"
         Splits(0)._ColumnProps(97)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(98)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(99)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(100)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(101)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(102)=   "Column(12)._ColStyle=139780"
         Splits(0)._ColumnProps(103)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(104)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(105)=   "Column(12).Order=13"
         Splits(1)._UserFlags=   0
         Splits(1).ExtendRightColumn=   -1  'True
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=13"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=139780"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(9)=   "Column(1).Width=1111"
         Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1032"
         Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=139777"
         Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(17)=   "Column(2).Width=8229"
         Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=8149"
         Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=139780"
         Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(25)=   "Column(2).DropDownList=1"
         Splits(1)._ColumnProps(26)=   "Column(3).Width=2725"
         Splits(1)._ColumnProps(27)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(28)=   "Column(3)._WidthInPix=2646"
         Splits(1)._ColumnProps(29)=   "Column(3).AllowSizing=0"
         Splits(1)._ColumnProps(30)=   "Column(3)._ColStyle=139780"
         Splits(1)._ColumnProps(31)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(32)=   "Column(3).AllowFocus=0"
         Splits(1)._ColumnProps(33)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(34)=   "Column(4).Width=2143"
         Splits(1)._ColumnProps(35)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(36)=   "Column(4)._WidthInPix=2064"
         Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=131588"
         Splits(1)._ColumnProps(38)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(39)=   "Column(4).AutoDropDown=1"
         Splits(1)._ColumnProps(40)=   "Column(4).DropDownList=1"
         Splits(1)._ColumnProps(41)=   "Column(5).Width=2725"
         Splits(1)._ColumnProps(42)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(5)._WidthInPix=2646"
         Splits(1)._ColumnProps(44)=   "Column(5).AllowSizing=0"
         Splits(1)._ColumnProps(45)=   "Column(5)._ColStyle=139780"
         Splits(1)._ColumnProps(46)=   "Column(5).Visible=0"
         Splits(1)._ColumnProps(47)=   "Column(5).AllowFocus=0"
         Splits(1)._ColumnProps(48)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(49)=   "Column(6).Width=1773"
         Splits(1)._ColumnProps(50)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(51)=   "Column(6)._WidthInPix=1693"
         Splits(1)._ColumnProps(52)=   "Column(6)._ColStyle=131588"
         Splits(1)._ColumnProps(53)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(54)=   "Column(6).AutoDropDown=1"
         Splits(1)._ColumnProps(55)=   "Column(6).DropDownList=1"
         Splits(1)._ColumnProps(56)=   "Column(7).Width=1244"
         Splits(1)._ColumnProps(57)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(58)=   "Column(7)._WidthInPix=1164"
         Splits(1)._ColumnProps(59)=   "Column(7).AllowSizing=0"
         Splits(1)._ColumnProps(60)=   "Column(7)._ColStyle=131588"
         Splits(1)._ColumnProps(61)=   "Column(7).Visible=0"
         Splits(1)._ColumnProps(62)=   "Column(7).AllowFocus=0"
         Splits(1)._ColumnProps(63)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(64)=   "Column(8).Width=3493"
         Splits(1)._ColumnProps(65)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(66)=   "Column(8)._WidthInPix=3413"
         Splits(1)._ColumnProps(67)=   "Column(8)._ColStyle=131588"
         Splits(1)._ColumnProps(68)=   "Column(8).Button=1"
         Splits(1)._ColumnProps(69)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(70)=   "Column(8).AutoDropDown=1"
         Splits(1)._ColumnProps(71)=   "Column(8).DropDownList=1"
         Splits(1)._ColumnProps(72)=   "Column(8).ButtonAlways=1"
         Splits(1)._ColumnProps(73)=   "Column(9).Width=2725"
         Splits(1)._ColumnProps(74)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(75)=   "Column(9)._WidthInPix=2646"
         Splits(1)._ColumnProps(76)=   "Column(9).AllowSizing=0"
         Splits(1)._ColumnProps(77)=   "Column(9)._ColStyle=139780"
         Splits(1)._ColumnProps(78)=   "Column(9).Visible=0"
         Splits(1)._ColumnProps(79)=   "Column(9).AllowFocus=0"
         Splits(1)._ColumnProps(80)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(81)=   "Column(10).Width=1826"
         Splits(1)._ColumnProps(82)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(83)=   "Column(10)._WidthInPix=1746"
         Splits(1)._ColumnProps(84)=   "Column(10)._ColStyle=131588"
         Splits(1)._ColumnProps(85)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(86)=   "Column(10).AutoDropDown=1"
         Splits(1)._ColumnProps(87)=   "Column(10).DropDownList=1"
         Splits(1)._ColumnProps(88)=   "Column(11).Width=2725"
         Splits(1)._ColumnProps(89)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(90)=   "Column(11)._WidthInPix=2646"
         Splits(1)._ColumnProps(91)=   "Column(11).AllowSizing=0"
         Splits(1)._ColumnProps(92)=   "Column(11)._ColStyle=139780"
         Splits(1)._ColumnProps(93)=   "Column(11).Visible=0"
         Splits(1)._ColumnProps(94)=   "Column(11).AllowFocus=0"
         Splits(1)._ColumnProps(95)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(96)=   "Column(12).Width=2963"
         Splits(1)._ColumnProps(97)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(98)=   "Column(12)._WidthInPix=2884"
         Splits(1)._ColumnProps(99)=   "Column(12)._ColStyle=131588"
         Splits(1)._ColumnProps(100)=   "Column(12).Button=1"
         Splits(1)._ColumnProps(101)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(102)=   "Column(12).ButtonAlways=1"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=130,.parent=21,.alignment=2,.valignment=2"
         _StyleDefs(38)  =   ":id=130,.bgcolor=&HCA570B&,.fgcolor=&HFFFFFF&,.locked=-1,.borderSize=1"
         _StyleDefs(39)  =   ":id=130,.borderColor=&HFFFFFF&"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=127,.parent=22"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=128,.parent=23"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=129,.parent=25"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=62,.parent=21,.alignment=0,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=22"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=23"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=25"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=66,.parent=21,.locked=0"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=22"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=23"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=25"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=122,.parent=21,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=22"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=23"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=25"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=114,.parent=21,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=22"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=23"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=25"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=106,.parent=21,.locked=-1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=103,.parent=22"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=104,.parent=23"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=105,.parent=25"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=74,.parent=21,.bgcolor=&HFFFFFF&"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=22"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=23"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=25"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=78,.parent=21,.bgcolor=&HFFFFFF&"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=22"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=23"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=25"
         _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=94,.parent=21,.bgcolor=&HFFFFFF&"
         _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=91,.parent=22"
         _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=92,.parent=23"
         _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=93,.parent=25"
         _StyleDefs(75)  =   "Splits(0).Columns(9).Style:id=16,.parent=21,.locked=-1"
         _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=13,.parent=22"
         _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=14,.parent=23"
         _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=15,.parent=25"
         _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=146,.parent=21,.locked=-1"
         _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=143,.parent=22"
         _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=144,.parent=23"
         _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=145,.parent=25"
         _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=98,.parent=21,.locked=-1"
         _StyleDefs(84)  =   "Splits(0).Columns(11).HeadingStyle:id=95,.parent=22"
         _StyleDefs(85)  =   "Splits(0).Columns(11).FooterStyle:id=96,.parent=23"
         _StyleDefs(86)  =   "Splits(0).Columns(11).EditorStyle:id=97,.parent=25"
         _StyleDefs(87)  =   "Splits(0).Columns(12).Style:id=138,.parent=21,.locked=-1"
         _StyleDefs(88)  =   "Splits(0).Columns(12).HeadingStyle:id=135,.parent=22"
         _StyleDefs(89)  =   "Splits(0).Columns(12).FooterStyle:id=136,.parent=23"
         _StyleDefs(90)  =   "Splits(0).Columns(12).EditorStyle:id=137,.parent=25"
         _StyleDefs(91)  =   "Splits(1).Style:id=79,.parent=1"
         _StyleDefs(92)  =   "Splits(1).CaptionStyle:id=88,.parent=4"
         _StyleDefs(93)  =   "Splits(1).HeadingStyle:id=80,.parent=2"
         _StyleDefs(94)  =   "Splits(1).FooterStyle:id=81,.parent=3"
         _StyleDefs(95)  =   "Splits(1).InactiveStyle:id=82,.parent=5"
         _StyleDefs(96)  =   "Splits(1).SelectedStyle:id=84,.parent=6"
         _StyleDefs(97)  =   "Splits(1).EditorStyle:id=83,.parent=7"
         _StyleDefs(98)  =   "Splits(1).HighlightRowStyle:id=85,.parent=8"
         _StyleDefs(99)  =   "Splits(1).EvenRowStyle:id=86,.parent=9"
         _StyleDefs(100) =   "Splits(1).OddRowStyle:id=87,.parent=10"
         _StyleDefs(101) =   "Splits(1).RecordSelectorStyle:id=89,.parent=11"
         _StyleDefs(102) =   "Splits(1).FilterBarStyle:id=90,.parent=12"
         _StyleDefs(103) =   "Splits(1).Columns(0).Style:id=134,.parent=79,.locked=-1"
         _StyleDefs(104) =   "Splits(1).Columns(0).HeadingStyle:id=131,.parent=80"
         _StyleDefs(105) =   "Splits(1).Columns(0).FooterStyle:id=132,.parent=81"
         _StyleDefs(106) =   "Splits(1).Columns(0).EditorStyle:id=133,.parent=83"
         _StyleDefs(107) =   "Splits(1).Columns(1).Style:id=32,.parent=79,.alignment=2,.locked=-1"
         _StyleDefs(108) =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=80"
         _StyleDefs(109) =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=81"
         _StyleDefs(110) =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=83"
         _StyleDefs(111) =   "Splits(1).Columns(2).Style:id=50,.parent=79,.locked=-1"
         _StyleDefs(112) =   "Splits(1).Columns(2).HeadingStyle:id=47,.parent=80"
         _StyleDefs(113) =   "Splits(1).Columns(2).FooterStyle:id=48,.parent=81"
         _StyleDefs(114) =   "Splits(1).Columns(2).EditorStyle:id=49,.parent=83"
         _StyleDefs(115) =   "Splits(1).Columns(3).Style:id=126,.parent=79,.bgcolor=&HFFFFFF&,.locked=-1"
         _StyleDefs(116) =   "Splits(1).Columns(3).HeadingStyle:id=123,.parent=80"
         _StyleDefs(117) =   "Splits(1).Columns(3).FooterStyle:id=124,.parent=81"
         _StyleDefs(118) =   "Splits(1).Columns(3).EditorStyle:id=125,.parent=83"
         _StyleDefs(119) =   "Splits(1).Columns(4).Style:id=118,.parent=79,.bgcolor=&HFFFFFF&"
         _StyleDefs(120) =   "Splits(1).Columns(4).HeadingStyle:id=115,.parent=80"
         _StyleDefs(121) =   "Splits(1).Columns(4).FooterStyle:id=116,.parent=81"
         _StyleDefs(122) =   "Splits(1).Columns(4).EditorStyle:id=117,.parent=83"
         _StyleDefs(123) =   "Splits(1).Columns(5).Style:id=110,.parent=79,.locked=-1"
         _StyleDefs(124) =   "Splits(1).Columns(5).HeadingStyle:id=107,.parent=80"
         _StyleDefs(125) =   "Splits(1).Columns(5).FooterStyle:id=108,.parent=81"
         _StyleDefs(126) =   "Splits(1).Columns(5).EditorStyle:id=109,.parent=83"
         _StyleDefs(127) =   "Splits(1).Columns(6).Style:id=20,.parent=79,.bgcolor=&HFFFFFF&"
         _StyleDefs(128) =   "Splits(1).Columns(6).HeadingStyle:id=17,.parent=80"
         _StyleDefs(129) =   "Splits(1).Columns(6).FooterStyle:id=18,.parent=81"
         _StyleDefs(130) =   "Splits(1).Columns(6).EditorStyle:id=19,.parent=83"
         _StyleDefs(131) =   "Splits(1).Columns(7).Style:id=54,.parent=79,.bgcolor=&HFFFFFF&"
         _StyleDefs(132) =   "Splits(1).Columns(7).HeadingStyle:id=51,.parent=80"
         _StyleDefs(133) =   "Splits(1).Columns(7).FooterStyle:id=52,.parent=81"
         _StyleDefs(134) =   "Splits(1).Columns(7).EditorStyle:id=53,.parent=83"
         _StyleDefs(135) =   "Splits(1).Columns(8).Style:id=58,.parent=79,.bgcolor=&HFFFFFF&,.locked=0"
         _StyleDefs(136) =   "Splits(1).Columns(8).HeadingStyle:id=55,.parent=80"
         _StyleDefs(137) =   "Splits(1).Columns(8).FooterStyle:id=56,.parent=81"
         _StyleDefs(138) =   "Splits(1).Columns(8).EditorStyle:id=57,.parent=83"
         _StyleDefs(139) =   "Splits(1).Columns(9).Style:id=70,.parent=79,.locked=-1"
         _StyleDefs(140) =   "Splits(1).Columns(9).HeadingStyle:id=67,.parent=80"
         _StyleDefs(141) =   "Splits(1).Columns(9).FooterStyle:id=68,.parent=81"
         _StyleDefs(142) =   "Splits(1).Columns(9).EditorStyle:id=69,.parent=83"
         _StyleDefs(143) =   "Splits(1).Columns(10).Style:id=150,.parent=79,.bgcolor=&HFFFFFF&"
         _StyleDefs(144) =   "Splits(1).Columns(10).HeadingStyle:id=147,.parent=80"
         _StyleDefs(145) =   "Splits(1).Columns(10).FooterStyle:id=148,.parent=81"
         _StyleDefs(146) =   "Splits(1).Columns(10).EditorStyle:id=149,.parent=83"
         _StyleDefs(147) =   "Splits(1).Columns(11).Style:id=102,.parent=79,.locked=-1"
         _StyleDefs(148) =   "Splits(1).Columns(11).HeadingStyle:id=99,.parent=80"
         _StyleDefs(149) =   "Splits(1).Columns(11).FooterStyle:id=100,.parent=81"
         _StyleDefs(150) =   "Splits(1).Columns(11).EditorStyle:id=101,.parent=83"
         _StyleDefs(151) =   "Splits(1).Columns(12).Style:id=142,.parent=79,.bgcolor=&HFFFFFF&"
         _StyleDefs(152) =   "Splits(1).Columns(12).HeadingStyle:id=139,.parent=80"
         _StyleDefs(153) =   "Splits(1).Columns(12).FooterStyle:id=140,.parent=81"
         _StyleDefs(154) =   "Splits(1).Columns(12).EditorStyle:id=141,.parent=83"
         _StyleDefs(155) =   "Named:id=33:Normal"
         _StyleDefs(156) =   ":id=33,.parent=0"
         _StyleDefs(157) =   "Named:id=34:Heading"
         _StyleDefs(158) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(159) =   ":id=34,.wraptext=-1"
         _StyleDefs(160) =   "Named:id=35:Footing"
         _StyleDefs(161) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(162) =   "Named:id=36:Selected"
         _StyleDefs(163) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(164) =   "Named:id=37:Caption"
         _StyleDefs(165) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(166) =   "Named:id=38:HighlightRow"
         _StyleDefs(167) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(168) =   "Named:id=39:EvenRow"
         _StyleDefs(169) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(170) =   "Named:id=40:OddRow"
         _StyleDefs(171) =   ":id=40,.parent=33"
         _StyleDefs(172) =   "Named:id=41:RecordSelector"
         _StyleDefs(173) =   ":id=41,.parent=34"
         _StyleDefs(174) =   "Named:id=42:FilterBar"
         _StyleDefs(175) =   ":id=42,.parent=33"
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   375
         Left            =   1350
         TabIndex        =   15
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   1305
         Width           =   1575
         Caption         =   " Listar"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManFlujoProceso.frx":1594
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGrabar 
         Height          =   375
         Left            =   3060
         TabIndex        =   14
         ToolTipText     =   " Graba la lista mostrada "
         Top             =   1305
         Width           =   1575
         Caption         =   " Grabar"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManFlujoProceso.frx":1B2E
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminaItem 
         Height          =   375
         Left            =   4815
         TabIndex        =   13
         ToolTipText     =   " Eliminar los datos de la cuenta selecccionada "
         Top             =   1305
         Width           =   1575
         Caption         =   " Eliminar Datos"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManFlujoProceso.frx":20C8
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminarTodo 
         Height          =   375
         Left            =   6525
         TabIndex        =   12
         ToolTipText     =   " Eliminar todas las cuentas de la lista "
         Top             =   1305
         Width           =   1575
         Caption         =   " Eliminar Todo"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManFlujoProceso.frx":2662
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   375
         Left            =   8235
         TabIndex        =   11
         Top             =   1305
         Width           =   1575
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManFlujoProceso.frx":2BFC
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton cmdVisibleImport 
         Height          =   375
         Left            =   5085
         TabIndex        =   10
         ToolTipText     =   " Importar configuración "
         Top             =   90
         Width           =   375
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "661;661"
         Value           =   "0"
         Picture         =   "frmManFlujoProceso.frx":3196
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "frmManFlujoProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrFlujo As New XArrayDB
Dim lArrDetalle(13) As Variant
Dim rsTipo As ADODB.Recordset
Dim rsActividad As ADODB.Recordset
Dim gsGrupo As String
Dim gsColumna As Integer

Dim VarX As Single
Dim VarY As Single
Dim cSepFormula As String

Dim Col_TipoCta As Integer, Col_Codigo As Integer, Col_Descripcion As Integer
Dim Col_CodActividad As Integer, Col_Actividad As Integer
Dim Col_CodTipoD As Integer, Col_TipoD As Integer, Col_FormulaD As Integer, Col_DetalleD As Integer
Dim Col_CodTipoH As Integer, Col_TipoH As Integer, Col_FormulaH As Integer, Col_DetalleH As Integer
Dim Col_Flag As Integer
Dim bInicio As Boolean
Const NUM_COL = 11

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub AsignaColumna()
    Col_TipoCta = 0
    Col_Codigo = 1
    Col_Descripcion = 2
    Col_CodActividad = 3
    Col_Actividad = 4
    
    Col_CodTipoD = 5
    Col_TipoD = 6
    Col_FormulaD = 7
    Col_DetalleD = 8
    
    Col_CodTipoH = 9
    Col_TipoH = 10
    Col_FormulaH = 11
    Col_DetalleH = 12
    
    Col_Flag = 13
End Sub

Private Sub cboMetodo_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdEliminaItem_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
            pSetFocus tdbcMes
       Exit Sub
    End If


    If IsNull(tdbgFlujo.Bookmark) Then
        Exit Sub
    End If
    

    If lArrFlujo(tdbgFlujo.Bookmark, Col_Flag) = "S" Then
        Mensajes "No se puede eliminar una cuenta de tipo titulo"
       Exit Sub
    End If


    cmdEliminaItem.Enabled = False
    DoEvents
    

    If MsgBox("Deseas eliminar los datos de la cuenta seleccionada", vbYesNo + vbQuestion) = vbYes Then
        lArrFlujo(tdbgFlujo.Bookmark, Col_CodActividad) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_Actividad) = Null
    
        lArrFlujo(tdbgFlujo.Bookmark, Col_CodTipoD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_TipoD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_FormulaD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_DetalleD) = Null
        
        lArrFlujo(tdbgFlujo.Bookmark, Col_CodTipoH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_TipoH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_FormulaH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_DetalleH) = Null
        
        tdbgFlujo.Refresh
       
        UpdateGrilla
    End If

    cmdEliminaItem.Enabled = True
End Sub

Private Sub cmdEliminarTodo_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo de proceso"
       Exit Sub
    End If

    If IsNull(tdbgFlujo.Bookmark) Then
        Exit Sub
    End If
    
    
    cmdEliminarTodo.Enabled = False
    DoEvents
    

    If MsgBox("Deseas eliminar los datos de la cuenta seleccionada", vbYesNo + vbQuestion) = vbYes Then
        Dim i As Integer
        For i = 0 To lArrFlujo.Count(1) - 1
        
        lArrFlujo(i, Col_CodActividad) = Null
        lArrFlujo(i, Col_Actividad) = Null
        
        lArrFlujo(i, Col_CodTipoD) = Null
        lArrFlujo(i, Col_TipoD) = Null
        lArrFlujo(i, Col_FormulaD) = Null
        lArrFlujo(i, Col_DetalleD) = Null
        
        lArrFlujo(i, Col_CodTipoH) = Null
        lArrFlujo(i, Col_TipoH) = Null
        lArrFlujo(i, Col_FormulaH) = Null
        lArrFlujo(i, Col_DetalleH) = Null
        
        
        Next i
    
       tdbgFlujo.Refresh
       
       UpdateGrilla
    End If

    cmdEliminarTodo.Enabled = True
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
       Exit Function
    End If
    
    If (lArrFlujo.Count(1) = 1 Or lArrFlujo.Count(2) = 1) And tdbgFlujo.Bookmark = 0 And CE(tdbgFlujo.Columns(Col_Codigo)) = "" Then
       ValidaCampos = True
       Exit Function
    End If
    
    Dim i As Integer
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If CE(lArrFlujo(i, Col_Actividad)) = "" And (CE(lArrFlujo(i, Col_DetalleD)) <> "" Or CE(lArrFlujo(i, Col_DetalleH)) <> "") Then
           Mensajes "Seleccione un tipo de Actividad, para la cuenta " & lArrFlujo(i, Col_Codigo)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = Col_Actividad
           pSetFocus tdbgFlujo
           Exit Function
        End If
    
        If CE(lArrFlujo(i, Col_TipoD)) = "" And CE(lArrFlujo(i, Col_DetalleD)) <> "" Then
           Mensajes "Verifique los datos del Ajuste Debe, para la cuenta " & lArrFlujo(i, Col_Codigo)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = Col_Actividad
           pSetFocus tdbgFlujo
           Exit Function
        End If
    
        If CE(lArrFlujo(i, Col_TipoH)) = "" And CE(lArrFlujo(i, Col_DetalleH)) <> "" Then
           Mensajes "Verifique los datos del Ajuste Haber, para la cuenta " & lArrFlujo(i, Col_Codigo)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = Col_Actividad
           pSetFocus tdbgFlujo
           Exit Function
        End If
    
        If CE(lArrFlujo(i, Col_TipoD)) <> "" And CE(lArrFlujo(i, Col_DetalleD)) = "" Then
           Mensajes "Ingrese una formula o valor para el Ajuste Debe, para la cuenta " & lArrFlujo(i, Col_Codigo)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = Col_DetalleD
           pSetFocus tdbgFlujo
           Exit Function
        End If
    
        If CE(lArrFlujo(i, Col_TipoH)) <> "" And CE(lArrFlujo(i, Col_DetalleH)) = "" Then
           Mensajes "Ingrese una formula o valor para el Ajuste Haber, para la cuenta " & lArrFlujo(i, Col_Codigo)
           tdbgFlujo.Bookmark = i
           tdbgFlujo.Col = Col_DetalleH
           pSetFocus tdbgFlujo
           Exit Function
        End If
    
    Next i
    
    ValidaCampos = True
End Function

Private Function CargaArregloDet(item As Integer) As Boolean
    CargaArregloDet = True
    
    lArrDetalle(0) = "INSERTAR"
    lArrDetalle(1) = gsEmpresa
    lArrDetalle(2) = gsAnio
    lArrDetalle(3) = tdbcMes.BoundText
    
    Dim sCuenta As String
    sCuenta = Replace(lArrFlujo(item, Col_TipoCta), Chr(13), "")
    
    If CE(sCuenta) = "CUENTA" Then
       lArrDetalle(4) = "CTA"
    Else
        lArrDetalle(4) = "ANA"
    End If
    
    lArrDetalle(5) = CE(lArrFlujo(item, Col_Codigo))  'cuenta
    lArrDetalle(6) = CE(lArrFlujo(item, Col_CodActividad))  'actividad
    '-----------------------------------------------------------------------
    lArrDetalle(7) = CE(lArrFlujo(item, Col_CodTipoD))  'tipo
    lArrDetalle(8) = CE(lArrFlujo(item, Col_FormulaD))  'formula
    
    If CE(lArrFlujo(item, Col_CodTipoD)) = "M" Then
       lArrDetalle(8) = CE(lArrFlujo(item, Col_DetalleD))  'formula
    End If
    
    lArrDetalle(9) = CE(lArrFlujo(item, Col_DetalleD))  'detalle
    '-----------------------------------------------------------------------
    lArrDetalle(10) = CE(lArrFlujo(item, Col_CodTipoH))  'tipo
    lArrDetalle(11) = CE(lArrFlujo(item, Col_FormulaH))  'formula
    
    If CE(lArrFlujo(item, Col_CodTipoH)) = "M" Then
       lArrDetalle(11) = CE(lArrFlujo(item, Col_DetalleH))  'formula
    End If
    
    lArrDetalle(12) = CE(lArrFlujo(item, Col_DetalleH))  'detalle
    '-----------------------------------------------------------------------
    lArrDetalle(13) = CE(cboMetodo.Text)  'METODO

End Function

Private Sub Grabar()
    UpdateGrilla
    
    If ValidaCampos = False Then Exit Sub

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas
    

    Dim lArrDet(13) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcMes.BoundText

    lArrDet(4) = ""
    lArrDet(5) = ""
    lArrDet(6) = ""
    lArrDet(7) = ""
    lArrDet(8) = ""
    lArrDet(9) = ""
    lArrDet(10) = ""
    lArrDet(11) = ""
    lArrDet(12) = ""
    
    lArrDet(13) = cboMetodo.Text
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoProceso", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    For i = 0 To lArrFlujo.Count(1) - 1
        If CE(lArrFlujo(i, Col_Actividad)) <> "" Then
            
                If CargaArregloDet(i) = True Then
                    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_FlujoProceso", lArrDetalle(), False) = False Then
                        Screen.MousePointer = vbNormal
                        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                        'tdbgFlujo.HighlightRowStyle = "HighlightRow"
                        tdbgFlujo.MarqueeStyle = dbgHighlightRow
                        tdbgFlujo.Bookmark = i
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
    

    If MsgBox("Desea importar los datos del mes seleccionado", vbYesNo + vbQuestion) = vbYes Then
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

'Private Sub cmdImportarFra_Click()
'    fraImportar.Visible = cmdImportarFra.Value
'End Sub

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
          "FROM TABLA WHERE TAB_CTABLA='070' AND EMP_CCODIGO='" & gsEmpresa & "' " & _
          "ORDER BY TAB_CDESCRIPCAMPO "
    
    Call CerrarRecordSet(rsTipo)
    Call LlenarRecordSet(sql, rsTipo)
    
    Set tdbdTipo.DataSource = rsTipo
    '--------------------------------------------------
    sql = "SELECT TAB_CDESCRIPCAMPO , TAB_CCODIGO " & _
          "FROM TABLA WHERE TAB_CTABLA='072' AND EMP_CCODIGO='" & gsEmpresa & "' " & _
          "ORDER BY TAB_CDESCRIPCAMPO "
    
    Call CerrarRecordSet(rsActividad)
    Call LlenarRecordSet(sql, rsActividad)
    
    
    Set tdbdActividad.DataSource = rsActividad
    
End Sub

Private Sub GeneraArreglo(Mes As String)
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion
    
    If tdbcMes.BoundText <> "" And cboMetodo.Text = "" Then
        Mensajes "Seleccione el metodo"
        pSetFocus cboMetodo
        Exit Sub
    End If
    
    If tdbcMes.Text <> "" And cboMetodo.Text <> "" Then
        
        sql = "spCn_FlujoProceso 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & Mes & "','','','','','','','','','','" & cboMetodo.Text & "'"
        
        
        If cboMetodo.Text = "DIRECTO" Then
            Call GridArreglo(lArrFlujo, tdbgFlujo, sql, "Col_TipoCtaRes='C'")
        Else
            Call GridArreglo(lArrFlujo, tdbgFlujo, sql)
        End If
        
        
        
        If lArrFlujo.Count(2) < NUM_COL Then
           lArrFlujo.ReDim 0, 0, 0, NUM_COL
        End If
    Else
        lArrFlujo.ReDim 0, 0, 0, NUM_COL
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
        If CE(lArrFlujo(i, Col_Codigo)) <> "" Then
           Contador = Contador + 1
        End If
    Next i
    
    CuentaFilas = Contador
End Function


Private Sub cmdRefresh_Click()
    If bInicio = False Then Exit Sub
    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    
    LlenaListas
    
    DoEvents
    
    GeneraArreglo tdbcMes.BoundText
    
    DoEvents
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
    
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
    bInicio = False
    DoEvents
    Call AsignaColumna
    
    Me.Height = 6525
    Me.Width = 11550
    

    tdbgFlujo.FetchRowStyle = True
    tdbgFlujo.Splits(0).MarqueeStyle = dbgHighlightRow
    
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
        tdbgFlujo.Splits(1).Columns(9).Button = False
        
        cmdVisibleImport.Enabled = False
    Else
        cmdGrabar.Enabled = True

        cmdEliminaItem.Enabled = True
        cmdEliminarTodo.Enabled = True
        cmdImportar.Enabled = True
        tdbgFlujo.Splits(0).Locked = False
        tdbgFlujo.Splits(0).Locked = True
        cmdVisibleImport.Enabled = True
        tdbgFlujo.Splits(1).Columns(9).Button = True
    End If
        
    cmdVisibleImport_Click
    
    DoEvents
    bInicio = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        fraTodo.Width = Me.Width
        fraTodo.Height = Me.Height
        'fraEstructura.Height = Me.Height - 2100
        'fraEstructura.Width = Me.Width - 150
        tdbgFlujo.Height = Me.Height - 2300
        tdbgFlujo.Width = Me.Width - 200
        tdbgFlujo.Splits(1).ScrollBars = dbgNone
        tdbgFlujo.Splits(1).ScrollBars = dbgAutomatic
        'fraLeyenda.Width = fraEstructura.Width
        'fraLeyenda.Top = fraEstructura.Top + fraEstructura.Height
    End If
    
    Exit Sub
    
serror:
End Sub


'Private Sub tdbcColumna_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        arbuAgregarVar_Click
'    End If
'
'End Sub

Private Sub tdbcMes_ItemChange()
    If tdbcMes.Text = "" Then
        tdbgFlujo.Splits(0).Locked = True
    Else
        tdbgFlujo.Splits(0).Locked = False
    End If
    
    cmdRefresh_Click
    
End Sub



Private Sub tdbdActividad_DropDownClose()
    tdbgFlujo.Columns(Col_CodActividad) = tdbdActividad.Columns(1).Value
    tdbgFlujo.Columns(Col_Actividad) = tdbdActividad.Columns(0).Value
    DoEvents
    On Error Resume Next
    tdbgFlujo.Update
    DoEvents
    pSetFocus tdbgFlujo

End Sub



Private Sub tdbdTipo_DropDownClose()
    If gsColumna = Col_TipoD Then
        tdbgFlujo.Columns(Col_CodTipoD) = tdbdTipo.Columns(1).Value
        tdbgFlujo.Columns(Col_TipoD) = tdbdTipo.Columns(0).Value
        'tdbgFlujo.Columns(Col_FormulaD) = ""
        'tdbgFlujo.Columns(Col_DetalleD) = ""
    End If
    
    If gsColumna = Col_TipoH Then
        tdbgFlujo.Columns(Col_CodTipoH) = tdbdTipo.Columns(1).Value
        tdbgFlujo.Columns(Col_TipoH) = tdbdTipo.Columns(0).Value
        'tdbgFlujo.Columns(Col_FormulaH) = ""
        'tdbgFlujo.Columns(Col_DetalleH) = ""
    End If
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
        If CE(lArrFlujo(tdbgFlujo.Bookmark, Col_Flag)) = "S" Then
           Cancel = 1
        End If
    End If
    
    If ColIndex = Col_DetalleD Then
        If tdbgFlujo.Columns(Col_CodTipoD) = "M" Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
    If ColIndex = Col_DetalleH Then
        If tdbgFlujo.Columns(Col_CodTipoH) = "M" Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub tdbgFlujo_ButtonClick(ByVal ColIndex As Integer)
    Dim Tipo As String
    
    If ColIndex = Col_DetalleD Then Tipo = CE(tdbgFlujo.Columns(Col_CodTipoD).Value)
    If ColIndex = Col_DetalleH Then Tipo = CE(tdbgFlujo.Columns(Col_CodTipoH).Value)
    
    If ColIndex = Col_DetalleD And Tipo <> "M" And Tipo <> "" Then
        frmFormulas.pFormula = CE(tdbgFlujo.Columns(Col_FormulaD).Value)
        frmFormulas.pObservacion = CE(tdbgFlujo.Columns(Col_DetalleD).Value)
        frmFormulas.pFormulario = Me.Name
        frmFormulas.pTipo = Tipo
        frmFormulas.pMetodo = cboMetodo.Text
        frmFormulas.pPeriodo = tdbcMes.BoundText
        frmFormulas.pAjuste = Col_FormulaD
        frmFormulas.Show vbModal
    End If
    
    If ColIndex = Col_DetalleH And Tipo <> "M" And Tipo <> "" Then
        frmFormulas.pFormula = CE(tdbgFlujo.Columns(Col_FormulaH).Value)
        frmFormulas.pObservacion = CE(tdbgFlujo.Columns(Col_DetalleH).Value)
        frmFormulas.pFormulario = Me.Name
        frmFormulas.pTipo = Tipo
        frmFormulas.pMetodo = cboMetodo.Text
        frmFormulas.pPeriodo = tdbcMes.BoundText
        frmFormulas.pAjuste = Col_FormulaH
        frmFormulas.Show vbModal
    End If
    
End Sub



Private Sub tdbgFlujo_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    'If lArrFlujo Is Nothing Or IsNull(tdbgFlujo.Bookmark) Then
    '    Exit Sub
    'End If
    
    On Error GoTo serror
    If Split = 0 Then
        If lArrFlujo(Bookmark, Col_Flag) = "1" Then
            RowStyle.BackColor = &HFFF9D7   'celeste
        End If
    
        If lArrFlujo(Bookmark, Col_Flag) = "2" Then
            RowStyle.BackColor = &HC4DBFF   'melon
        End If
    
        If lArrFlujo(Bookmark, Col_Flag) = "3" Then
            RowStyle.BackColor = &HCEFFF8   'amarillo
        End If
    
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
    If gsColumna = 9 Then
       If KeyCode <> 13 And KeyCode <> vbKeyEnd And KeyCode <> vbKeyHome And _
          KeyCode <> 46 And KeyCode <> vbKeyF2 And KeyCode <> vbKeyBack And _
          KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyUp And _
          KeyCode <> vbKeyDown Then
            If Not IsNumeric(Chr(KeyCode)) Then
               KeyCode = 0
            End If
       End If
    End If

    If gsColumna = Col_TipoD And KeyCode = vbKeyDelete Then
        lArrFlujo(tdbgFlujo.Bookmark, Col_CodTipoD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_TipoD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_FormulaD) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_DetalleD) = Null
        
        tdbgFlujo.Refresh
    End If
    
    If gsColumna = Col_TipoH And KeyCode = vbKeyDelete Then
        lArrFlujo(tdbgFlujo.Bookmark, Col_CodTipoH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_TipoH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_FormulaH) = Null
        lArrFlujo(tdbgFlujo.Bookmark, Col_DetalleH) = Null
        
        tdbgFlujo.Refresh
    End If
End Sub

Private Sub tdbgFlujo_KeyPress(KeyAscii As Integer)
    If gsColumna = 9 Then
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
    
    If KeyAscii = 13 And gsColumna >= Col_DetalleH Then
       On Error GoTo serror
       UpdateGrilla
       DoEvents
   
       tdbgFlujo.Col = Col_Actividad
       tdbgFlujo.Bookmark = tdbgFlujo.Bookmark + 1
       
serror:
       pSetFocus tdbgFlujo
    
    End If
    
End Sub

Private Sub tdbgFlujo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gsColumna = tdbgFlujo.Col
End Sub


