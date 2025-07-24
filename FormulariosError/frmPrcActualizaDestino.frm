VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcActualizaDestino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproceso de Asiento de Destino"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   Icon            =   "frmPrcActualizaDestino.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   8475
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "RECALCULA LAS CUENTAS DE DESTINO"
      TabPicture(0)   =   "frmPrcActualizaDestino.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "RECONSTRUYE LAS CUENTAS DE DESTINO"
      TabPicture(1)   =   "frmPrcActualizaDestino.frx":0EE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   2400
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   3780
         Begin TrueOleDBList70.TDBCombo tdbcMesAux 
            Height          =   300
            Left            =   180
            TabIndex        =   10
            Top             =   720
            Width           =   3345
            _ExtentX        =   5900
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
            _PropDict       =   $"frmPrcActualizaDestino.frx":0F02
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
         Begin MSComctlLib.ProgressBar pgbAvanceAux 
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   1440
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   344
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblMesAux 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   270
            TabIndex        =   13
            Top             =   1080
            Width           =   3255
         End
         Begin MSForms.CommandButton cmdProcesarAux 
            Height          =   435
            Left            =   1215
            TabIndex        =   12
            Top             =   1845
            Width           =   1665
            Caption         =   "Procesar"
            PicturePosition =   327683
            Size            =   "2937;767"
            Picture         =   "frmPrcActualizaDestino.frx":0F89
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2280
         Index           =   0
         Left            =   -72720
         TabIndex        =   1
         Top             =   2160
         Width           =   3780
         Begin VB.CheckBox chkMes 
            Caption         =   "Hasta el mes seleccionado"
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   3300
         End
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   180
            TabIndex        =   3
            Top             =   360
            Width           =   3345
            _ExtentX        =   5900
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
            _PropDict       =   $"frmPrcActualizaDestino.frx":291B
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
         Begin MSComctlLib.ProgressBar pgbAvance 
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   1200
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   344
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSForms.CommandButton cmdProcesar 
            Height          =   435
            Left            =   990
            TabIndex        =   6
            Top             =   1560
            Width           =   1665
            Caption         =   "Procesar"
            PicturePosition =   327683
            Size            =   "2937;767"
            Picture         =   "frmPrcActualizaDestino.frx":29A2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label lblMes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   270
            TabIndex        =   5
            Top             =   840
            Width           =   3255
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmPrcActualizaDestino.frx":4334
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   825
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   7545
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NO REGISTRAR ningun movimiento  del mes seleccionado al ejecutar el Proceso."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   7545
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"frmPrcActualizaDestino.frx":44A6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   825
         Index           =   1
         Left            =   -74640
         TabIndex        =   8
         Top             =   1200
         Width           =   7545
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NO REGISTRAR ningun movimiento  del mes seleccionado al ejecutar el Proceso."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Index           =   0
         Left            =   -74640
         TabIndex        =   7
         Top             =   600
         Width           =   7545
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
      TabIndex        =   17
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcActualizaDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsMante As clsMantoTablas
Dim gsGrupo As String
Public gsMensaje As Boolean
Public gsSinSaldos As Boolean
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Function ProcesoSinActSaldos(Mes As String) As Boolean
    
    Dim lArrMnt() As Variant
    ReDim lArrMnt(4) As Variant
    On Local Error GoTo ErrorEjecucion
    lArrMnt(0) = gsEmpresa          ' Empresa
    lArrMnt(1) = gsAnio             ' Codigo
    lArrMnt(2) = Mes                ' Nombre
    lArrMnt(3) = "A"                ' Nombre Plantilla
    lArrMnt(4) = gsUsuario          ' Usuario
    If CierreMes(Mes) = False Then
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoDestinoV2", lArrMnt(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            ProcesoSinActSaldos = False
            Exit Function
        End If
        
    End If
    ProcesoSinActSaldos = True
    Exit Function
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Private Function Proceso(Mes As String) As Boolean
        
    Dim lArrMnt() As Variant
    ReDim lArrMnt(4) As Variant
    On Local Error GoTo ErrorEjecucion
    Me.SSTab1.Tab = 0
    DoEvents
    lArrMnt(0) = gsEmpresa          ' Empresa
    lArrMnt(1) = gsAnio             ' Codigo
    lArrMnt(2) = Mes                ' Nombre
    lArrMnt(3) = "A"                ' Nombre Plantilla
    lArrMnt(4) = gsUsuario          ' Usuario
    If CierreMes(Mes) = False Then
    
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoDestino", lArrMnt(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Proceso = False
            Exit Function
        End If
        
    End If
    Proceso = True
    Exit Function
    
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    
End Function


Public Sub cmdProcesar_Click()
Dim sql As String
Dim CodLib As String
Dim NomMes As String

If tdbcMes.BoundText = "00" Then
    CodLib = "01"
    NomMes = "ENERO"
ElseIf tdbcMes.BoundText = "13" Or tdbcMes.BoundText = "14" Then
    CodLib = "12"
    NomMes = "DICIEMBRE"
Else
    CodLib = tdbcMes.BoundText
    NomMes = tdbcMes.Text
End If

    cmdProcesar.Enabled = False
    DoEvents
    
        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & CodLib & "' and Lib_cTipoLibro = '03' and Estado ='A'"
        If ExisteDato(sql) = True Then
            Mensajes "No se puede procesar las Cuentas de Destino, debido a que el Libro Electrónico del periodo de " + NomMes + " del " + gsAnio + " se encuentra generado." & Salto(2) & "* En caso de no haber realizado su envío proceda a eliminarlo." & Salto(1) & "* En caso de haberlo enviado, Desbloquee el libro y corrija manualmente su información.", vbInformation
                cmdProcesar.Enabled = True
                Exit Sub
        End If
        
    Call Procesar
    cmdProcesar.Enabled = True
End Sub

Public Sub Cerrar()
    Unload Me
End Sub

Public Sub Procesar()
    Dim i As Integer
    Dim Mes As String
    Dim Ultimo As Integer
    Dim inicio As Integer
    Dim Retorno As Boolean
    On Error GoTo serror
    
    lblMES.Caption = "INICIANDO ..."
    Call EscribirLog("Iniciando la generacion de cuentas de destino de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    Screen.MousePointer = vbHourglass
    
    Ultimo = Val(tdbcMes.BoundText)
    
    pgbAvance.Min = 0
    If Ultimo = 0 Then
        pgbAvance.Max = 1
    Else
        pgbAvance.Max = Ultimo
    End If
    
    pgbAvance.Value = 0
    
    If chkMes.Value = vbUnchecked Then
        inicio = Ultimo
    Else
        inicio = 0
    End If
    
    Set clsMante = New clsMantoTablas
    clsMante.InicializaClase
    clsMante.BeginTrans
    
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        lblMES.Caption = "PROCES. " & NombreMes(Mes)
        lblMES.Refresh
        pgbAvance.Value = i
        pgbAvance.Refresh
        DoEvents
        
        If gsSinSaldos = False Then
            Retorno = Proceso(Mes)
        Else
            Retorno = ProcesoSinActSaldos(Mes)
        End If
        
        If Retorno = False Then Exit For
    Next i
    
    If Retorno = True Then
        clsMante.CommitTrans
        clsMante.FinalizaClase
    Else
        clsMante.FinalizaClase
    End If
    
    Set clsMante = Nothing
    
    lblMES.Caption = "PROCESO TERMINADO ..."
    
    Call EscribirLog("Finalizo la generacion de cuentas de destino de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    If gsMensaje = True Then Mensajes "Proceso ha terminado con exito", vbInformation + vbOKOnly
    Screen.MousePointer = vbDefault
    
    Exit Sub
serror:
    Call EscribirLog("error de actualizacion de cuentas de destino, [" & Err.Description & "]  de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Sub

Public Sub ProcesarAux()
    Dim i As Integer
    Dim Mes As String

    Screen.MousePointer = vbHourglass
    lblMesAux.Caption = "PROCESO INICIADO..."
    Call EscribirLog("Iniciando la replica de cuentas desde el mes seleccionado de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    pgbAvanceAux.Value = 0
    pgbAvanceAux.Min = 0
    pgbAvanceAux.Max = 10
    Set clsMante = New clsMantoTablas
    
    Dim lArrMnt() As Variant
    ReDim lArrMnt(3) As Variant
    On Local Error GoTo ErrorEjecucion
    lArrMnt(0) = "PROCESO_REPLMES01"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = tdbcMesAux.BoundText

    If CierreMes(Mes) = False Then
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Set clsMante = Nothing
            Exit Sub
        End If
        
    End If
    
    
    Set clsMante = Nothing
    
    lblMesAux.Caption = "PROCESO TERMINADO ..."
    pgbAvanceAux.Value = 10
    
    Call EscribirLog("Finalizo la replica de cuentas desde el mes seleccionado de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Mensajes "Proceso ha terminado con exito", vbInformation + vbOKOnly
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrorEjecucion:
    Mensajes Err.Description
    Call EscribirLog("error de actualizacion de cuentas de destino desde el mes seleecionado, [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Sub


Private Sub cmdProcesarAux_Click()
Dim sql As String
Dim CodLib As String
Dim NomMes As String

If tdbcMesAux.BoundText = "00" Then
    CodLib = "01"
    NomMes = "ENERO"
ElseIf tdbcMesAux.BoundText = "13" Or tdbcMesAux.BoundText = "14" Then
    CodLib = "12"
    NomMes = "DICIEMBRE"
Else
    CodLib = tdbcMesAux.BoundText
    NomMes = tdbcMesAux.Text
End If

    cmdProcesarAux.Enabled = False
    DoEvents

        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & CodLib & "' and Lib_cTipoLibro = '03' and Estado ='A'"
        If ExisteDato(sql) = True Then
            Mensajes "No se puede procesar las Cuentas de Destino, debido a que el Libro Electrónico del periodo de " + NomMes + " del " + gsAnio + " se encuentra generado." & Salto(2) & "* En caso de no haber realizado su envío proceda a eliminarlo." & Salto(1) & "* En caso de haberlo enviado, Desbloquee el libro y corrija manualmente su información.", vbInformation
                cmdProcesarAux.Enabled = True
                Exit Sub
        End If
    
    Call ProcesarAux
    cmdProcesarAux.Enabled = True
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    gsMensaje = True
    Call Centrar_form(Me)
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    Call LlenaComboMesAddItem(tdbcMesAux)
    
    tdbcMes.ReBind
    tdbcMes.BoundText = gsPeriodo
    
    Dim cMes As String
    cMes = gsPeriodo
    If cMes = "00" Then cMes = "01"
    If cMes > "12" Then cMes = "12"
    
    tdbcMesAux.ReBind
    tdbcMesAux.BoundText = cMes
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdProcesar.Enabled = False
    Else
        Me.cmdProcesar.Enabled = True
    End If
    Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(SSTab1, Me)
        Call CentrarTitulo(lblTitulo, SSTab1, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub


