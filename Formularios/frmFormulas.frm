VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFormulas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diseñador de Formulas"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmFormulas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuerpo 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.TextBox tdbtFormula 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1260
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   7215
      End
      Begin VB.Frame fraAnalisis 
         Height          =   735
         Left            =   90
         TabIndex        =   8
         Top             =   945
         Width           =   10050
         Begin VB.ComboBox tdbcAjuste 
            Height          =   315
            Left            =   6030
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   225
            Width           =   2385
         End
         Begin TrueOleDBList70.TDBCombo tdbcAnalisis 
            Height          =   300
            Left            =   1125
            TabIndex        =   9
            Tag             =   "enabled"
            Top             =   225
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   12806
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
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=10980"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=10901"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            CellTips        =   2
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
            MaxComboItems   =   10
            AddItemSeparator=   ";"
            _PropDict       =   $"frmFormulas.frx":0ECA
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Analisis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   13
            Top             =   270
            Width           =   675
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Ajuste"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   5265
            TabIndex        =   12
            Top             =   270
            Width           =   540
         End
         Begin MSForms.CommandButton arbuAgregarAna 
            Height          =   390
            Left            =   8550
            TabIndex        =   11
            ToolTipText     =   "Agregar cuenta de analisis a la formula"
            Top             =   225
            Width           =   1305
            Caption         =   " Agregar"
            PicturePosition =   327683
            Size            =   "2302;688"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraVariables 
         Height          =   825
         Left            =   90
         TabIndex        =   14
         Top             =   1620
         Width           =   10050
         Begin VB.ComboBox tdbcColumna 
            Height          =   315
            Left            =   6030
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox tdbtCuenta 
            Height          =   330
            Left            =   1170
            TabIndex        =   15
            Top             =   270
            Width           =   1230
         End
         Begin TrueOleDBList70.TDBCombo tdbcVariables 
            Height          =   300
            Left            =   1170
            TabIndex        =   17
            Tag             =   "enabled"
            Top             =   270
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   12700
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
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=10980"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=10901"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            CellTips        =   2
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
            MaxComboItems   =   10
            AddItemSeparator=   ";"
            _PropDict       =   $"frmFormulas.frx":0F51
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Columna"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   5220
            TabIndex        =   21
            Top             =   315
            Width           =   750
         End
         Begin MSForms.CommandButton arbuAgregarVar 
            Height          =   390
            Left            =   8550
            TabIndex        =   20
            ToolTipText     =   "Agregar cuenta contable a la formula"
            Top             =   225
            Width           =   1305
            Caption         =   " Agregar"
            PicturePosition =   327683
            Size            =   "2302;688"
            Picture         =   "frmFormulas.frx":0FD8
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   315
            Width           =   600
         End
         Begin VB.Label lblDescripcion 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2475
            TabIndex        =   18
            Top             =   270
            Width           =   5865
         End
      End
      Begin VB.Frame fraValor 
         Height          =   780
         Left            =   5130
         TabIndex        =   26
         Top             =   2385
         Width           =   5010
         Begin TDBNumber6Ctl.TDBNumber tdbnValor2 
            Height          =   315
            Left            =   1680
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   556
            Calculator      =   "frmFormulas.frx":1572
            Caption         =   "frmFormulas.frx":1592
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmFormulas.frx":15FE
            Keys            =   "frmFormulas.frx":161C
            Spin            =   "frmFormulas.frx":1666
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   16711680
            Format          =   "###,###,###,##0.00"
            HighlightText   =   0
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
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnValor 
            Height          =   315
            Left            =   960
            TabIndex        =   30
            Top             =   240
            Width           =   2370
            _Version        =   65536
            _ExtentX        =   4180
            _ExtentY        =   556
            Calculator      =   "frmFormulas.frx":168E
            Caption         =   "frmFormulas.frx":16AE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmFormulas.frx":171A
            Keys            =   "frmFormulas.frx":1738
            Spin            =   "frmFormulas.frx":1782
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.00"
            HighlightText   =   0
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   225
            TabIndex        =   29
            Top             =   315
            Width           =   450
         End
         Begin MSForms.CommandButton arbuAgregarVal 
            Height          =   390
            Left            =   3510
            TabIndex        =   28
            ToolTipText     =   "Agregar valor numerico a la formula"
            Top             =   225
            Width           =   1305
            Caption         =   " Agregar"
            PicturePosition =   327683
            Size            =   "2302;688"
            Picture         =   "frmFormulas.frx":17AA
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraOperador 
         Height          =   780
         Left            =   90
         TabIndex        =   22
         Top             =   2385
         Width           =   5010
         Begin VB.ComboBox tdbcOperador 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   270
            Width           =   2205
         End
         Begin MSForms.CommandButton arbuAgregarOpe 
            Height          =   390
            Left            =   3510
            TabIndex        =   25
            ToolTipText     =   "Agregar operador a la formula"
            Top             =   225
            Width           =   1305
            Caption         =   " Agregar"
            PicturePosition =   327683
            Size            =   "2302;688"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Operador"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   24
            Top             =   315
            Width           =   810
         End
      End
      Begin VB.TextBox tdbtObserva 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   870
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3240
         Width           =   7170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fórmula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label12 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   4
         Top             =   3285
         Width           =   1020
      End
      Begin MSForms.CommandButton arbuBorrar 
         Height          =   390
         Left            =   8640
         TabIndex        =   3
         ToolTipText     =   "Limpiar formula y observacion"
         Top             =   270
         Width           =   1305
         Caption         =   " Limpiar"
         PicturePosition =   327683
         Size            =   "2302;688"
         Picture         =   "frmFormulas.frx":1D44
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.CommandButton cmdCancelar 
      Height          =   390
      Left            =   5265
      TabIndex        =   6
      Top             =   4365
      Width           =   1305
      Caption         =   " Cancelar"
      PicturePosition =   327683
      Size            =   "2302;688"
      Picture         =   "frmFormulas.frx":22DE
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAceptar 
      Height          =   390
      Left            =   3690
      TabIndex        =   7
      Top             =   4365
      Width           =   1305
      Caption         =   " Aceptar"
      PicturePosition =   327683
      Size            =   "2302;688"
      Picture         =   "frmFormulas.frx":2878
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmFormulas
'    Project    : Contabilidad
'
'    Description: Formulario de mantenimiento de formulas
'--------------------------------------------------------------------------------
Dim cSepFormula As String
Public pTipo As String
Public pPeriodo As String
Public pFormula As String
Public pObservacion As String
Public pFormulario As String
Public pMetodo As String
Public pAjuste As Integer
Public pColumna As Integer
Dim nContador As Integer

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       arbuAgregarAna_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de agregar analisis
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub arbuAgregarAna_Click()
Dim cadena As String
Dim Cuenta As String
Dim Tipo As String
Tipo = pTipo

If CE(tdbcAnalisis.Text) = "" Then
    Mensajes "Seleccione un tipo de analisis de la lista"
    pSetFocus tdbcAnalisis
    Exit Sub
End If

Cuenta = tdbcAnalisis.Columns(0)
cadena = tdbcAnalisis.Columns(1)


If Cuenta <> "" Then
    Dim Columna As String
    Columna = ""
    
    If ValidaInsOper(4) = True Then
        If Tipo = "H" Then Columna = Left(tdbcAjuste.List(tdbcAjuste.ListIndex), 2)
        
        tdbtFormula.SelText = cSepFormula + "ANA" + Cuenta + Columna + cSepFormula
        tdbtFormula = Trim(tdbtFormula)
        tdbtFormula.SelStart = Len(tdbtFormula)
        pSetFocus tdbtFormula
        
        If Tipo = "H" Then Columna = " (" & Columna & ") "
        
        tdbtObserva.SelText = cSepFormula + "[ " + cadena + Columna + " ]" + cSepFormula
        tdbtObserva = Trim(tdbtObserva)
        tdbtObserva.SelStart = Len(tdbtObserva)
        pSetFocus tdbcOperador
        
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       arbuAgregarOpe_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de agregar operacion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub arbuAgregarOpe_Click()
Dim cOper As String
cOper = Trim(Left(tdbcOperador.Text, 1))
' Verifica que el operador no este vacío
If cOper <> "" Then
    If ValidaInsOper(TipoOperador(cOper)) = True Then
        tdbtFormula.SelText = cSepFormula + cOper + cSepFormula
        tdbtFormula = Trim(tdbtFormula)
        tdbtFormula.SelStart = Len(tdbtFormula)
        pSetFocus tdbtFormula
        
        If cOper = "/" Then
            tdbtObserva.SelText = cSepFormula + vbCrLf + Replicar("-", 90) + vbCrLf + cSepFormula
            tdbtObserva = Trim(tdbtObserva)
            tdbtObserva.SelStart = Len(tdbtObserva)
            
            If pTipo = "H" Then
                pSetFocus tdbcVariables
            Else
                tdbtCuenta.SelStart = 0
                tdbtCuenta.SelLength = Len(tdbtCuenta.Text)
                
                pSetFocus tdbtCuenta
            End If
        Else
            tdbtObserva.SelText = cSepFormula + cOper + cSepFormula
            tdbtObserva = Trim(tdbtObserva)
            tdbtObserva.SelStart = Len(tdbtObserva)
            
            If pTipo = "H" Then
                pSetFocus tdbcVariables
            Else
                tdbtCuenta.SelStart = 0
                tdbtCuenta.SelLength = Len(tdbtCuenta.Text)
            
                pSetFocus tdbtCuenta
            End If
            
        End If
        
    End If
End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       arbuAgregarVal_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de agregar valor
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub arbuAgregarVal_Click()
    If CE(tdbnValor.Text) <> "" Then
        If ValidaInsOper(4) = True Then
            tdbtFormula.SelText = cSepFormula + Trim(Str(tdbnValor.Value)) + cSepFormula
            tdbtFormula = Trim(tdbtFormula)
            tdbtFormula.SelStart = Len(tdbtFormula)
            pSetFocus tdbtFormula
            
            tdbtObserva.SelText = cSepFormula + Trim(Str(tdbnValor.Value)) + cSepFormula
            tdbtObserva = Trim(tdbtObserva)
            tdbtObserva.SelStart = Len(tdbtObserva)
            pSetFocus tdbtObserva
            
        End If
    Else
        Mensajes "Ingrese un valor numérico"
        pSetFocus tdbnValor
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       arbuAgregarVar_Click
' Description:       Evento que se ejecuta al hacer clic en el botn agregarvariable
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub arbuAgregarVar_Click()
Dim cadena As String
Dim Cuenta As String
Dim Tipo As String
Tipo = pTipo

If Tipo = "H" Then
    If CE(tdbcVariables.Text) = "" Then
        Mensajes "Seleccione una cuenta de la lista"
        pSetFocus tdbcVariables
        Exit Sub
    End If

   Cuenta = tdbcVariables.Columns(0)
   cadena = tdbcVariables.Columns(1)
Else

    If CE(tdbtCuenta.Text) = "" Then
        Mensajes "Ingrese una cuenta valida"
            pSetFocus tdbtCuenta
        Exit Sub
    End If

   Cuenta = tdbtCuenta.Text
   cadena = lblDescripcion.Caption
End If

If Cuenta <> "" Then
    Dim Columna As String
    Columna = ""
    
    If ValidaInsOper(4) = True Then
        If Tipo = "H" Then Columna = Left(tdbcColumna.List(tdbcColumna.ListIndex), 2)
        
        tdbtFormula.SelText = cSepFormula + "CTA" + Cuenta + Columna + cSepFormula
        tdbtFormula = Trim(tdbtFormula)
        tdbtFormula.SelStart = Len(tdbtFormula)
        pSetFocus tdbtFormula
        
        If Tipo = "H" Then Columna = " (" & Columna & ") "
        
        tdbtObserva.SelText = cSepFormula + "[ " + cadena + Columna + " ]" + cSepFormula
        tdbtObserva = Trim(tdbtObserva)
        tdbtObserva.SelStart = Len(tdbtObserva)
        pSetFocus tdbcOperador
        
    End If
End If

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       arbuBorrar_Click
' Description:       Evento que se ejecuta al hacer clic en el boton borrar
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub arbuBorrar_Click()
    tdbtFormula.Text = ""
    tdbtObserva.Text = ""

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaCombos
' Description:       Procedimiento de llenado de datos enlos combos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaCombos()
    Call combo_operador
    Call combo_columna
    Call combo_ajuste
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       combo_ajuste
' Description:       Procedimiento de llenado de combo de ajuste
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub combo_ajuste()
    tdbcAjuste.Clear
    If pFormulario = "frmManFlujoReporte" Then
        tdbcAjuste.List(0) = "AD" + " - " + "Ajustes Debe"
        tdbcAjuste.List(1) = "AH" + " - " + "Ajustes Haber"
        tdbcAjuste.List(2) = "SA" + " - " + "Saldo"
        
    Else
        tdbcAjuste.List(0) = "AD" + " - " + "Ajustes Debe"
        tdbcAjuste.List(1) = "AH" + " - " + "Ajustes Haber"
    End If
    
    On Error Resume Next
    tdbcAjuste.ListIndex = 0

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       combo_columna
' Description:       Procedimiento de llenado de combo de columna
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub combo_columna()
    tdbcColumna.Clear
    If pFormulario = "frmManFlujoReporte" Then
        tdbcColumna.List(0) = "AD" + " - " + "Ajustes Debe"
        tdbcColumna.List(1) = "AH" + " - " + "Ajustes Haber"
        tdbcColumna.List(2) = "SA" + " - " + "Saldo"
    Else
        tdbcColumna.List(0) = "PI" + " - " + "Periodo Inicial"
        tdbcColumna.List(1) = "PF" + " - " + "Periodo Final"
        tdbcColumna.List(2) = "VD" + " - " + "Variacion Debe"
        tdbcColumna.List(3) = "VH" + " - " + "Variacion Haber"
        tdbcColumna.List(4) = "AD" + " - " + "Ajustes Debe"
        tdbcColumna.List(5) = "AH" + " - " + "Ajustes Haber"
    End If
    
    On Error Resume Next
    tdbcColumna.ListIndex = 0

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       combo_operador
' Description:       Procedimiento de llenado de combo de operador
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub combo_operador()
    tdbcOperador.List(0) = "(" + " ; " + "Abrir parentesis"
    tdbcOperador.List(1) = "(" + " ; " + "Abrir parentesis"
    tdbcOperador.List(2) = ")" + " ; " + "Cerrar parentesis"
    tdbcOperador.List(3) = "+" + " ; " + "Sumar"
    tdbcOperador.List(4) = "-" + " ; " + "Restar"
    tdbcOperador.List(5) = "*" + " ; " + "Multiplicar"
    
    On Error Resume Next
    tdbcOperador.ListIndex = 3
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaVariables
' Description:       procedimient dellenado de combo de variables
'
' Parameters :       Tipo (String)
'--------------------------------------------------------------------------------
Private Sub CargaVariables(Tipo As String)
    Dim sqlSp As String
    Dim condicion As String
    
    If pMetodo = "INDIRECTO" Then
       condicion = ""
    Else
        condicion = " AND PLA_CCUENTACONTABLE<'60' "
    End If
    
    If Tipo = "H" Then
                
        sqlSp = "spCn_FlujoSaldos 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & gsAnio & "', '" & pPeriodo & "'"
        
        Dim arr As New XArrayDB
        
        Call ComboArreglo(arr, tdbcVariables, sqlSp, "PLA_CCUENTACONTABLE<>'XX'" & condicion)
    
    End If
    
    If pMetodo = "INDIRECTO" Then
        sqlSp = "select TAB_CCODIGO, TAB_CDESCRIPCAMPO " & _
                "FROM TABLA " & _
                "WHERE Emp_cCodigo= '" & gsEmpresa & "' and TAB_CTABLA='073' " & _
                "ORDER BY TAB_CCODIGO"
                
        LlenarComboAddItem tdbcAnalisis, sqlSp
    End If
    On Error Resume Next
    tdbcVariables.Bookmark = 0
    tdbcAnalisis.Bookmark = 0
    DoEvents
    
    tdbcVariables.ReBind
    tdbcAnalisis.ReBind
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdAceptar_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de aceptar
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdAceptar_Click()
    If ValidaFormula(tdbtFormula.Text, 8) = True Then
       

       If pFormulario = "frmManFlujoProceso" Then
          frmManFlujoProceso.tdbgFlujo.Columns(pAjuste) = tdbtFormula.Text
          frmManFlujoProceso.tdbgFlujo.Columns(pAjuste + 1) = tdbtObserva.Text
          frmManFlujoProceso.UpdateGrilla
       End If
       
       If pFormulario = "frmManFlujoReporte" Then
          frmManFlujoReporte.tdbgFlujo.Columns(4) = tdbtFormula.Text
          frmManFlujoReporte.tdbgFlujo.Columns(5) = tdbtObserva.Text
          frmManFlujoReporte.UpdateGrilla
       End If
       
       If pFormulario = "frmManPatrimonioNeto" Then
          frmManPatrimonioNeto.grdPatrimonio.Columns(pColumna) = tdbtFormula.Text
          frmManPatrimonioNeto.UpdateGrilla
       End If
       
       
       Unload Me
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdCancelar_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de cancelar
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdCancelar_Click()
    nContador = 1
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Activate
' Description:       Evento que se ejecuta al activarse el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Activate()
    
    If pFormulario = "frmManPatrimonioNeto" And nContador = 1 Then
       Me.tdbtFormula.Text = pFormula
       pSetFocus Me.tdbtCuenta
    End If
    nContador = nContador + 1
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el formulario
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call cmdCancelar_Click
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    lblDescripcion.Caption = ""
    DoEvents
    cSepFormula = " "
    
    Screen.MousePointer = vbHourglass
    CargaVariables (pTipo)
    DoEvents

    Me.tdbtFormula.Text = pFormula
    
    
    
    Me.tdbtObserva.Text = pObservacion

    tdbtFormula.SelStart = Len(tdbtFormula.Text)
    tdbtObserva.SelStart = Len(tdbtObserva.Text)

    Screen.MousePointer = vbNormal
    
    If pMetodo = "INDIRECTO" Then
        fraAnalisis.Visible = True
        tdbtFormula.Height = 700
        
    ElseIf pMetodo = "DIRECTO" Then
        tdbtFormula.Height = 1395
        fraAnalisis.Visible = False
        
    ElseIf pMetodo = "PATRIMONIO" Then
        fraCuerpo.Height = 3240
        cmdAceptar.Top = 3300
        cmdCancelar.Top = 3300
        Me.Height = 4100 + 300
        
        tdbtFormula.Height = 1395
        fraAnalisis.Visible = False
        
    End If
    
    If pTipo = "C" Or pTipo = "D" Or pTipo = "B" Then
        tdbcColumna.Visible = False
        Label14(1).Visible = False
        tdbcVariables.Visible = False
        tdbtCuenta.Visible = True
        lblDescripcion.Visible = True
        
        fraAnalisis.Visible = False
    Else 'Hoja trabajo
        tdbcColumna.Visible = True
        Label14(1).Visible = True
        tdbcVariables.Visible = True
        tdbtCuenta.Visible = False
        lblDescripcion.Visible = False
    End If
    
    If pFormulario = "frmManFlujoReporte" Then
        fraValor.Visible = False
    Else
        fraValor.Visible = True
    End If
    
    Call LlenaCombos
    DoEvents
    pSetFocus tdbcVariables
    

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcAjuste_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el combo de ajuste
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcAjuste_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarAna_Click
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcAnalisis_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en elcombo de analisis
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcAnalisis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If tdbcAjuste.Visible Then
            pSetFocus tdbcAjuste
        End If
    End If

End Sub

'Private Sub ImgCerrar_Click()
'
'    Unload Me
'End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcColumna_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el combo de columna
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcColumna_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarVar_Click
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcOperador_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en elcombo de operador
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcOperador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarOpe_Click
    End If

End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcVariables_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el combo de variables
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcVariables_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If tdbcColumna.Visible Then
            pSetFocus tdbcColumna
        Else
            arbuAgregarVar_Click
        End If
    End If

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ValidaInsOper
' Description:       Procedimiento de validacion al insertar un operador
'
' Parameters :       nCurrTOper (Integer)
'--------------------------------------------------------------------------------
Private Function ValidaInsOper(nCurrTOper As Integer) As Boolean
Dim nLastTOper As Integer, nNextTOper As Integer
Dim i As Integer, nAbre As Integer, nCierra As Integer
ValidaInsOper = False
tdbtFormula.SelLength = IIf(tdbtFormula.SelStart = 0, 0, 1)
' Verifica que la inserción se haga al inicio o final o en un espacio en blanco
If tdbtFormula.SelText = cSepFormula Or tdbtFormula.SelStart = 0 Or tdbtFormula.SelStart = Len(tdbtFormula) Then
    nLastTOper = TipoOperador(RetornaOperadorFormula(True))
    nNextTOper = TipoOperador(RetornaOperadorFormula(False))
    Select Case nCurrTOper
        Case Is = 4
            ' NO puede ir despues de los tipos de operador
            If nLastTOper = 4 Or nLastTOper = 2 Then Exit Function
            ' NO Puede ir antes de los tipos de operador
            If nNextTOper = 4 Or nNextTOper = 1 Then Exit Function
        Case Is = 1 ' Operador de Agrupacion de Apertura
            ' NO puede ir despues de los tipos de operador
            If nLastTOper = 4 Or nLastTOper = 2 Then Exit Function
            ' NO Puede ir antes de los tipos de operador
            If nNextTOper = 2 Or nNextTOper = 3 Then Exit Function
        Case Is = 2  ' Operador de Agrupacion de Cierre
            ' NO puede ir despues de los tipos de operador
            If nLastTOper = 1 Or nLastTOper = 3 Then Exit Function
            ' NO Puede ir antes de los tipos de operador
            If nNextTOper = 4 Or nNextTOper = 1 Then Exit Function
            nAbre = 0
            nCierra = 0
            For i = 1 To Len(tdbtFormula)
                If Mid(tdbtFormula, i, 1) = Left(tdbcOperador.List(1), 1) Then nAbre = nAbre + 1
                If Mid(tdbtFormula, i, 1) = Left(tdbcOperador.List(2), 1) Then nCierra = nCierra + 1
            Next
            ' Si no hay tantos operadores de agrupacion abiertos
            If nAbre < nCierra + 1 Then Exit Function
        Case Is = 3 ' Si es operador aritmético
            ' La formula no puede empezar con operador aritmético
            If Len(tdbtFormula) = 0 Then Exit Function
            ' NO puede ir despues de los tipos de operador
            If nLastTOper = 1 Or nLastTOper = 3 Then Exit Function
            ' NO Puede ir antes de los tipos de operador
            If nNextTOper = 2 Or nNextTOper = 3 Then Exit Function
    End Select
    ValidaInsOper = True
End If
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       TipoOperador
' Description:       Funcion que valida el tipo de operador
'
' Parameters :       sFormula (String)
'--------------------------------------------------------------------------------
Private Function TipoOperador(sFormula As String) As Integer
Dim i As Integer
sFormula = Trim(sFormula)
If sFormula = "" Then TipoOperador = 0: Exit Function
For i = 1 To Me.tdbcOperador.ListCount
    If sFormula = Left(tdbcOperador.List(i), 1) Then
        If i <= 2 Then TipoOperador = i: Exit Function  ' Operadores de Agrupación
         TipoOperador = 3   ' Operador Aritmético
         Exit Function
    End If
Next
TipoOperador = 4
End Function
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       RetornaOperadorFormula
' Description:       Funcion que retornael operador de la formula
'
' Parameters :       bLast (Boolean = False)
'--------------------------------------------------------------------------------
Private Function RetornaOperadorFormula(Optional bLast As Boolean = False) As String
Dim nPosIni As Integer, nPosFin As Integer
RetornaOperadorFormula = ""
If Len(tdbtFormula) = 0 Then Exit Function
' Que devuelva el Operador Anterior
If bLast = True Then
    ' Si el Punto de Inserción esta al INICIO
    If tdbtFormula.SelStart = 0 Then Exit Function
    '***
    nPosFin = tdbtFormula.SelStart
    nPosIni = InStrRev(tdbtFormula, cSepFormula, nPosFin)
    If nPosIni = 0 Then RetornaOperadorFormula = Left(tdbtFormula, nPosFin): Exit Function
    RetornaOperadorFormula = Trim(Mid(tdbtFormula, nPosIni, nPosFin - nPosIni + 1))
' Que devuelva el Operador Siguiente
Else
    ' Si el Punto de Inserción esta al FINAL
    If tdbtFormula.SelStart = Len(tdbtFormula) Then Exit Function
    '***
    nPosIni = IIf(tdbtFormula.SelStart < 1, 1, tdbtFormula.SelStart + 2)
    nPosFin = InStr(nPosIni, tdbtFormula, cSepFormula)
    If nPosFin = 0 Then RetornaOperadorFormula = Mid(tdbtFormula, nPosIni): Exit Function
    RetornaOperadorFormula = Trim(Mid(tdbtFormula, nPosIni, nPosFin - nPosIni))
End If
End Function


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Replicar
' Description:       Funcion dereplica de valores segun el parametro
'
' Parameters :       cadena (String)
'                    veces (Integer)
'--------------------------------------------------------------------------------
Public Function Replicar(cadena As String, veces As Integer) As String
    Dim i As Integer
    Replicar = ""
    For i = 1 To veces
        Replicar = Replicar & cadena
    Next i
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       RecibirDatos
' Description:       Procedimiento de recibir los datos del formulario de busquedas
'
' Parameters :       lControl (String)
'                    param0 (String)
'                    param1 (String)
'                    param2 (String)
'--------------------------------------------------------------------------------
Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)

    tdbtCuenta.Text = Trim(param0)
    lblDescripcion.Caption = Trim(param1)
    Unload frmBuscador
    pSetFocus tdbtCuenta

    DoEvents
   
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbnValor_KeyDown
' Description:       Evento que se ejecuta al presionar una teclaen el campo de valor
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbnValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarVal_Click
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCuenta_Change
' Description:       Evento que se ejecuta al cmabiar el codigo de cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCuenta_Change()
    If CE(tdbtCuenta.Text) = "" Then lblDescripcion.Caption = ""
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCuenta_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el campo de cuenta
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        ' CUENTAS DE DETALLE Y DE TITULO = CuentasFilt
        Call LlamaBuscar(frmBuscador, Me.tdbtCuenta.Name, Me.tdbtCuenta.Name, "CuentasFilt", Me, gsPeriodo, Me.tdbtCuenta.Text)
    End If
    
    If KeyCode = 13 Then
        If tdbtCuenta.Text <> "" And Me.Enabled = True Then
            lblDescripcion.Caption = ExisteCta(tdbtCuenta.Text)
            If lblDescripcion.Caption = "" Then
                Mensajes "Cuenta no existe"
                tdbtCuenta.Text = ""
                pSetFocus tdbtCuenta
            End If
        End If

       arbuAgregarVar_Click
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCuenta_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque en el campo de cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCuenta_LostFocus()
    If tdbtCuenta.Text <> "" And Me.Enabled = True Then
        lblDescripcion.Caption = ExisteCta(tdbtCuenta.Text)
        If lblDescripcion.Caption = "" Then
           Mensajes "Cuenta no existe"
           tdbtCuenta.Text = ""
           pSetFocus tdbtCuenta
        End If
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtFormula_Change
' Description:       Evento que se ejecuta al cambiar la formual
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtFormula_Change()
    If CE(tdbtFormula.Text) = "" Then
       arbuBorrar_Click
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtFormula_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el campo de formula
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtFormula_KeyPress(KeyAscii As Integer)
Dim nPos As Integer

If KeyAscii = 8 And CE(tdbtFormula.Text) <> "" Then
' Presiona BACKSPACE
    nPos = InStrRev(tdbtFormula, cSepFormula, Len(tdbtFormula))
    If nPos >= 0 Then tdbtFormula = Trim(Left(tdbtFormula, nPos))
    tdbtFormula.SelStart = Len(tdbtFormula)
End If

End Sub


