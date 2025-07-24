VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcCierreMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo de meses (Seguridad)"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   Icon            =   "frmPrcCierreMes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   6195
   Begin VB.Frame fraTodo 
      Height          =   6900
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   6090
      Begin VB.Frame Frame1 
         Height          =   5565
         Left            =   135
         TabIndex        =   7
         Top             =   675
         Width           =   5790
         Begin TrueOleDBList70.TDBCombo TDBLibro 
            Height          =   300
            Left            =   2805
            TabIndex        =   2
            Top             =   1010
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
            _PropDict       =   $"frmPrcCierreMes.frx":0ECA
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
         Begin VB.Frame fraHastaMes 
            Height          =   855
            Left            =   225
            TabIndex        =   8
            Top             =   105
            Width           =   5325
            Begin VB.CheckBox chkTodos 
               Caption         =   "Todos"
               Height          =   255
               Left            =   2295
               TabIndex        =   13
               Top             =   555
               Width           =   825
            End
            Begin VB.OptionButton optCerrar 
               Caption         =   "BLOQUEAR"
               Height          =   315
               Left            =   1170
               TabIndex        =   0
               Top             =   225
               Value           =   -1  'True
               Width           =   1320
            End
            Begin VB.OptionButton optAbrir 
               Caption         =   "DESBLOQUEAR"
               Height          =   360
               Left            =   2955
               TabIndex        =   1
               Top             =   225
               Width           =   1575
            End
            Begin VB.CheckBox chkHastaMes 
               Caption         =   "Hasta el mes seleccionado"
               Height          =   285
               Left            =   2445
               TabIndex        =   9
               Top             =   2925
               Width           =   2355
            End
         End
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   2805
            TabIndex        =   3
            Top             =   1365
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
            _PropDict       =   $"frmPrcCierreMes.frx":0F51
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
         Begin TrueOleDBList70.TDBList tdblEstadoMeses 
            Height          =   3585
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   6324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PERIODO"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "MES"
            Columns(1).DataField=   "Nombre Cuenta"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "ESTADO"
            Columns(2).DataField=   "%"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=3704"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3625"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=688"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=609"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   1
            BorderStyle     =   1
            MatchEntry      =   0
            RightToLeft     =   0   'False
            MatchCompare    =   -6
            MatchCol        =   0
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   5
            MultiSelect     =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            ExposeCellMode  =   0
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutUrl       =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            ListField       =   ""
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            RowMember       =   ""
            MouseIcon       =   0
            MouseIcon.vt    =   3
            MousePointer    =   0
            MatchEntryTimeout=   2000
            AnimateWindow   =   0
            AnimateWindowDirection=   0
            AnimateWindowTime=   200
            AnimateWindowClose=   0
            DataView        =   0
            GroupByCaption  =   "Drag a column header here to group by that column"
            ScrollTrack     =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            AddItemSeparator=   ";"
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=248,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&"
            _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            Left            =   1230
            TabIndex        =   14
            Top             =   1010
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mes de Proceso"
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
            Left            =   1230
            TabIndex        =   10
            Top             =   1395
            Width           =   1335
         End
      End
      Begin MSForms.CommandButton cmdProcesar 
         Height          =   390
         Left            =   1215
         TabIndex        =   4
         Top             =   6345
         Width           =   1665
         Caption         =   " Procesar"
         PicturePosition =   327683
         Size            =   "2937;688"
         Picture         =   "frmPrcCierreMes.frx":0FD8
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   390
         Left            =   3015
         TabIndex        =   5
         Top             =   6345
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;688"
         Picture         =   "frmPrcCierreMes.frx":1572
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Este proceso permite bloquear los meses seleccionados para que no se puedan realizar cambios por el usuario."
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   5550
      End
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
      TabIndex        =   12
      Top             =   0
      Width           =   4125
   End
End
Attribute VB_Name = "frmPrcCierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdProcesar_Click()
    If gsAdmin <> "1" Then
        Mensajes "El usuario no es un administrador, no puede abrir/cerrar meses"
        Exit Sub
    End If
    
    Dim cMesIni As String
    Dim cMesFin As String
    Dim cMes As String
    Dim i As Integer
    
    cMesFin = tdbcMes.BoundText
    
    If chkHastaMes.Value = vbChecked Then
        cMesIni = "00"
    Else
        cMesIni = cMesFin
    End If
    
    'Seleccion del Check Todos, hara un barrido desde el mes "00" hasta el "14"
    If Me.chkTodos.Value = 1 Then
        cMes = "00"
        cMesIni = "00"
        cMesFin = "14"
    End If
    
    For i = cMesIni To cMesFin
        cMes = Right("00" & CE(i), 2)
        Call GrabarCierre(cMes)
    Next i
    
    Call LlenaEstadoMes
    
    DoEvents
    
    Mensajes "Proceso terminado."
End Sub

Private Function GrabarCierre(cPeriodo As String) As Boolean
    ' *** Generar el cierre
    Dim clsMante As clsMantoTablas
    Dim lArrMnt() As Variant
    Dim cMensaje As String
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    Screen.MousePointer = vbHourglass
    On Local Error GoTo ErrorEjecucion
    ReDim lArrMnt(6) As Variant
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = cPeriodo
    lArrMnt(4) = "A"
    lArrMnt(5) = gsUsuario
    lArrMnt(6) = TDBLibro.BoundText
    
    lArrMnt(0) = "EXISTEREG"     ' Accion
    
    If optCerrar.Value = False Then
        lArrMnt(4) = "A"
    Else
        lArrMnt(4) = "I"
    End If
    
    'PGBV - 02012013
    Dim sql As String

'    sql = "SELECT * FROM CNT_CIERRE WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & tdbcMes.BoundText & "' and TipoLibro = '" & TDBLibro.BoundText & "'"
sql = "SELECT * FROM CNT_CIERRE WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & cPeriodo & "' and TipoLibro = '" & TDBLibro.BoundText & "'"
    If ExisteDato(sql) = True Then
        lArrMnt(0) = "EDITAR"     ' Accion
    Else
        If optCerrar.Value = True Then
            lArrMnt(0) = "INSERTAR"     ' Accion
        End If
    End If

    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If chkHastaMes.Value = vbChecked Then
        cMensaje = "HASTA "
    Else
        cMensaje = ""
    End If
    
    
    Screen.MousePointer = vbDefault
    Exit Function
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
If gsPeriodo = "" Then
        tdbcMes.BoundText = "00"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
   
    Call LlenaComboMesApeAddItem(tdbcMes)
    If gsPeriodo = "" Then
        tdbcMes.BoundText = "00"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
    Call Llenarlibros

    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdProcesar.Enabled = False
    Else
        Me.cmdProcesar.Enabled = True
    End If

    Call LlenaEstadoMes
End Sub

Private Sub Llenarlibros()
    TDBLibro.AddItem "01" & ";" & "APERTURA"
    TDBLibro.AddItem "02" & ";" & "CAJA INGRESOS"
    TDBLibro.AddItem "03" & ";" & "DIARIO"
    TDBLibro.AddItem "04" & ";" & "CAJA EGRESOS"
    TDBLibro.AddItem "05" & ";" & "VENTAS"
    TDBLibro.AddItem "06" & ";" & "COMPRAS"
    TDBLibro.AddItem "07" & ";" & "DIFERENCIA DE CAMBIO"
    TDBLibro.AddItem "08" & ";" & "CIERRE"
    TDBLibro.Bookmark = 0
    TDBLibro.ListField = "column1"
    TDBLibro.BoundColumn = "column0"
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fratodo, Me)
        Call CentrarTitulo(lblTitulo, fratodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub LlenaEstadoMes()

    Dim sqlSp  As String
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    Dim i As Integer
    
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    
    sqlSp = "spCn_GrabaCierre 'BUSCAESTADO', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "','', '', '" & TDBLibro.BoundText & "'"
    arrDatos = Array(sqlSp)
    
    tdblEstadoMeses.Clear
    
    Call CerrarRecordSet(rsArreglo)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        Do While Not rsArreglo.EOF
            tdblEstadoMeses.AddItem rsArreglo!Per_cPeriodo & ";" & rsArreglo!PER_CDESCRIPPERIODO & ";" & rsArreglo!Estado
            rsArreglo.MoveNext
        Loop
    End If
    
    Set clDatos = Nothing
    Call CerrarRecordSet(rsArreglo)
    Set rsArreglo = Nothing
End Sub

Private Sub tdblEstadoMeses_Click()
tdbcMes.BoundText = tdblEstadoMeses.Columns(0).Value
'MsgBox Me.tdblEstadoMeses.Columns("Nombre Cuenta").Value
End Sub

Private Sub TDBLibro_ItemChange()
Call LlenaEstadoMes
End Sub
