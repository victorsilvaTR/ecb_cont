VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{280CBA51-58D5-46E7-950F-21C424D176BF}#1.0#0"; "RmFrame.ocx"
Begin VB.Form frmManEntidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Entidades"
   ClientHeight    =   6828
   ClientLeft      =   2400
   ClientTop       =   2040
   ClientWidth     =   10848
   Icon            =   "frmManEntidades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6828
   ScaleWidth      =   10848
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   615
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   4500
      Begin VB.Frame Frame4 
         Height          =   1905
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4365
         Begin VB.OptionButton optCodigo 
            Caption         =   "Por Codigo"
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
            Left            =   1620
            TabIndex        =   22
            Top             =   810
            Width           =   1425
         End
         Begin VB.OptionButton optNombre 
            Caption         =   "Por Nombre"
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
            Left            =   1620
            TabIndex        =   21
            Top             =   510
            Width           =   1380
         End
         Begin VB.OptionButton optRuc 
            Caption         =   "Por Ruc"
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
            Left            =   1620
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1065
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   435
            Left            =   390
            TabIndex        =   25
            Top             =   1260
            Width           =   1665
            Caption         =   " Imprimir"
            PicturePosition =   327683
            Size            =   "2937;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton Command1 
            Height          =   435
            Left            =   2220
            TabIndex        =   23
            Top             =   1290
            Width           =   1665
            Caption         =   " Salir"
            PicturePosition =   327683
            Size            =   "2937;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin pRmFrame.RmFrame rmBarra 
         Height          =   330
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4470
         _ExtentX        =   7895
         _ExtentY        =   572
         BorderStyle     =   6
         BorderType      =   16384
         Caption         =   " Imprimir Reporte Ordenado por   "
         CaptionAlign    =   36
         BackColor       =   11035138
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor1  =   11035138
         GradientColor2  =   16306595
         BackgroundType  =   1
         Begin VB.Image ImgCerrar 
            Height          =   210
            Left            =   4170
            Picture         =   "frmManEntidades.frx":0ECA
            Stretch         =   -1  'True
            Top             =   60
            Width           =   240
         End
      End
   End
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   10575
      _ExtentX        =   18648
      _ExtentY        =   11028
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consultas de Entidades"
      TabPicture(0)   =   "frmManEntidades.frx":11F1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Entidades"
      TabPicture(1)   =   "frmManEntidades.frx":120D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74640
         TabIndex        =   10
         Top             =   480
         Width           =   9915
         Begin TabDlg.SSTab SSTab1 
            Height          =   1335
            Left            =   360
            TabIndex        =   12
            Top             =   480
            Width           =   8310
            _ExtentX        =   14669
            _ExtentY        =   2350
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Codigo y Nombres"
            TabPicture(0)   =   "frmManEntidades.frx":1229
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label2(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label2(5)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "tdbcEntidadBus"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "tdbtDescripcionBus"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "tdbtCodigoBus"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "chkTipo"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "Otros Datos"
            TabPicture(1)   =   "frmManEntidades.frx":1245
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label2(10)"
            Tab(1).Control(1)=   "Label2(11)"
            Tab(1).Control(2)=   "tdbtNroDocBus"
            Tab(1).Control(3)=   "tdbtDireccionBus"
            Tab(1).ControlCount=   4
            Begin VB.CheckBox chkTipo 
               Caption         =   "Por Tipo"
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
               Left            =   3120
               TabIndex        =   1
               Top             =   480
               Width           =   1095
            End
            Begin TDBText6Ctl.TDBText tdbtCodigoBus 
               Height          =   315
               Left            =   1200
               TabIndex        =   0
               Top             =   480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManEntidades.frx":1261
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":12CD
               Key             =   "frmManEntidades.frx":12EB
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "9"
               FormatMode      =   0
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   5
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
               Height          =   315
               Left            =   1200
               TabIndex        =   3
               Top             =   840
               Width           =   6570
               _Version        =   65536
               _ExtentX        =   11589
               _ExtentY        =   556
               Caption         =   "frmManEntidades.frx":133D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":13A9
               Key             =   "frmManEntidades.frx":13C7
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   120
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtDireccionBus 
               Height          =   315
               Left            =   -73320
               TabIndex        =   6
               Top             =   840
               Width           =   4335
               _Version        =   65536
               _ExtentX        =   7646
               _ExtentY        =   556
               Caption         =   "frmManEntidades.frx":1419
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1485
               Key             =   "frmManEntidades.frx":14A3
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   120
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtNroDocBus 
               Height          =   315
               Left            =   -73320
               TabIndex        =   5
               Top             =   480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManEntidades.frx":14F5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1561
               Key             =   "frmManEntidades.frx":157F
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   120
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TrueOleDBList70.TDBCombo tdbcEntidadBus 
               Height          =   300
               Left            =   4320
               TabIndex        =   2
               Top             =   480
               Width           =   3465
               _ExtentX        =   6117
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
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=826"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=762"
               Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=1122"
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
               EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
               LayoutName      =   ""
               LayoutFileName  =   ""
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTips        =   0
               EditHeight      =   276.095
               AutoSize        =   -1  'True
               GapHeight       =   36.283
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
               _PropDict       =   $"frmManEntidades.frx":15D1
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nro Documento"
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
               Index           =   11
               Left            =   -74760
               TabIndex        =   16
               Top             =   480
               Width           =   1305
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Direccion"
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
               Index           =   10
               Left            =   -74760
               TabIndex        =   15
               Top             =   840
               Width           =   780
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nombre"
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
               Index           =   5
               Left            =   240
               TabIndex        =   14
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Codigo"
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
               Index           =   4
               Left            =   240
               TabIndex        =   13
               Top             =   480
               Width           =   600
            End
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   2535
            Left            =   360
            TabIndex        =   4
            Top             =   1920
            Width           =   9315
            _ExtentX        =   16425
            _ExtentY        =   4466
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Cod Tipo"
            Columns(0).DataField=   "Ten_cTipoEntidad"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo"
            Columns(1).DataField=   "Ten_cNombreEntidad"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Codigo"
            Columns(2).DataField=   "Ent_cCodEntidad"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Razon Social / Nombres"
            Columns(3).DataField=   "Ent_cPersona"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Direccion"
            Columns(4).DataField=   "Ent_cDireccion"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tipo Doc."
            Columns(5).DataField=   "TipoDocumento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nro Documento"
            Columns(6).DataField=   "Ent_nRuc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=699"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=614"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2201"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1715"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1630"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=529"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=6414"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=6329"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2731"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(31)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=2286"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2201"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=532"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(39)=   "Column(6).Width=2815"
            Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2731"
            Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=532"
            Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.4,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.4,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            EmptyRows       =   -1  'True
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H80000008&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=Arial"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=Arial"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H800000&"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(26)  =   ":id=13,.fgcolor=&H80000008&"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filtrar Datos"
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
            Index           =   6
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   10155
         Begin VB.CheckBox chkPorcentajeSunat 
            Caption         =   "5% Sunat?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   63
            Top             =   4560
            Width           =   2175
         End
         Begin TrueOleDBList70.TDBCombo tdbcVinculoEconomico 
            Height          =   300
            Left            =   2040
            TabIndex        =   60
            Top             =   5040
            Width           =   6975
            _ExtentX        =   12298
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
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=868"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=804"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2709"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   2
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   5
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   276.095
            AutoSize        =   -1  'True
            GapHeight       =   36.283
            ListField       =   ""
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
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            AddItemSeparator=   ";"
            _PropDict       =   $"frmManEntidades.frx":1658
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
         Begin TrueOleDBList70.TDBCombo tdbcPais 
            Height          =   300
            Left            =   2040
            TabIndex        =   59
            Top             =   4560
            Width           =   3495
            _ExtentX        =   6160
            _ExtentY        =   529
            _LayoutType     =   0
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
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1651"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1651"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   2
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   5
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   276.095
            AutoSize        =   -1  'True
            GapHeight       =   36.283
            ListField       =   ""
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
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            AddItemSeparator=   ";"
            _PropDict       =   $"frmManEntidades.frx":16DF
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
         Begin VB.CheckBox chkActivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo"
            Enabled         =   0   'False
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
            Left            =   600
            TabIndex        =   51
            Tag             =   "_"
            Top             =   6195
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Frame fradescripcion 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   360
            TabIndex        =   41
            Top             =   3120
            Width           =   9735
            Begin VB.Frame fraconvenio 
               BorderStyle     =   0  'None
               Height          =   975
               Left            =   5400
               TabIndex        =   42
               Top             =   360
               Width           =   4335
               Begin VB.OptionButton OptNO 
                  Caption         =   "NO"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   58
                  Top             =   240
                  Width           =   855
               End
               Begin VB.OptionButton OptSi 
                  Caption         =   "SI"
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   56
                  Top             =   240
                  Width           =   495
               End
               Begin VB.CheckBox chkAconvenio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Aplica Convenio"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   36
                  Top             =   600
                  Width           =   1695
               End
               Begin TrueOleDBList70.TDBCombo tdbcAplicaconvenio 
                  Height          =   300
                  Left            =   2040
                  TabIndex        =   38
                  Top             =   600
                  Width           =   2265
                  _ExtentX        =   4001
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
                  Columns.Count   =   3
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=3"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=826"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=762"
                  Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(11)=   "Column(2).Width=1122"
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
                  EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  EditHeight      =   299.906
                  AutoSize        =   0   'False
                  GapHeight       =   36.283
                  ListField       =   "Ten_cNombreEntidad"
                  BoundColumn     =   "Ten_cTipoEntidad"
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
                  _PropDict       =   $"frmManEntidades.frx":1766
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
               Begin VB.Label Label1 
                  Caption         =   "Domiciliado"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   0
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1455
               End
            End
            Begin TDBText6Ctl.TDBText tdbtDireccion 
               Height          =   315
               Left            =   1680
               TabIndex        =   31
               Tag             =   "_"
               Top             =   60
               Width           =   5895
               _Version        =   65536
               _ExtentX        =   10398
               _ExtentY        =   556
               Caption         =   "frmManEntidades.frx":17ED
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1859
               Key             =   "frmManEntidades.frx":1877
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   0
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   120
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtNroDocumento 
               Height          =   300
               Left            =   1680
               TabIndex        =   33
               Tag             =   "_"
               Top             =   990
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Caption         =   "frmManEntidades.frx":18C9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1935
               Key             =   "frmManEntidades.frx":1953
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   0
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   15
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   -1
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TrueOleDBList70.TDBCombo tdbcTipoDocumento 
               Height          =   300
               Left            =   1680
               TabIndex        =   32
               Top             =   510
               Width           =   3465
               _ExtentX        =   6117
               _ExtentY        =   529
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               _DropdownWidth  =   6710
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Gru_cGrupo"
               Columns(0).DataField=   "Edoc_cTipoDoc"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Gru_cDescripLarga"
               Columns(1).DataField=   "TipDoc"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Tamanio"
               Columns(2).DataField=   "Tamanio"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=720"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=656"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2709"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=2709"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(15)=   "Column(2).AllowSizing=0"
               Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
               Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
               Splits.Count    =   1
               Appearance      =   1
               BorderStyle     =   1
               ComboStyle      =   2
               AutoCompletion  =   0   'False
               LimitToList     =   0   'False
               ColumnHeaders   =   0   'False
               ColumnFooters   =   0   'False
               DataMode        =   0
               DefColWidth     =   0
               Enabled         =   -1  'True
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               Caption         =   ""
               EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
               LayoutName      =   ""
               LayoutFileName  =   ""
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTips        =   0
               EditHeight      =   276.095
               AutoSize        =   -1  'True
               GapHeight       =   36.283
               ListField       =   "Gru_cDescripLarga"
               BoundColumn     =   "Gru_cGrupo"
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
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               MaxComboItems   =   10
               AddItemSeparator=   ";"
               _PropDict       =   $"frmManEntidades.frx":19A5
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=1,.bold=0,.fontsize=825,.italic=0"
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
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            Begin VB.Label lbldireccion 
               Caption         =   "Direccion"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nro Documento"
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
               Index           =   7
               Left            =   120
               TabIndex        =   44
               Top             =   1020
               Width           =   1305
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Doc"
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
               Left            =   120
               TabIndex        =   43
               Top             =   540
               Width           =   735
            End
         End
         Begin VB.Frame fradatos 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   360
            TabIndex        =   35
            Top             =   1560
            Width           =   7815
            Begin TDBText6Ctl.TDBText tdbtApaterno 
               Height          =   375
               Left            =   1680
               TabIndex        =   27
               Top             =   120
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   661
               Caption         =   "frmManEntidades.frx":1A2C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1A98
               Key             =   "frmManEntidades.frx":1AB6
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   40
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtAmaterno 
               Height          =   375
               Left            =   5520
               TabIndex        =   28
               Top             =   120
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
               _ExtentY        =   661
               Caption         =   "frmManEntidades.frx":1AFA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1B66
               Key             =   "frmManEntidades.frx":1B84
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   40
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText tdbtNombres 
               Height          =   375
               Left            =   1680
               TabIndex        =   29
               Top             =   720
               Width           =   5895
               _Version        =   65536
               _ExtentX        =   10398
               _ExtentY        =   661
               Caption         =   "frmManEntidades.frx":1BC8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEntidades.frx":1C34
               Key             =   "frmManEntidades.frx":1C52
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   1
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   0
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   "a"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   40
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin VB.Label LblApa 
               Caption         =   "Apellido Paterno"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblAma 
               Caption         =   "Apellido Materno"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   39
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblnombres 
               Caption         =   "Nombres"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   720
               Width           =   1215
            End
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   2040
            TabIndex        =   30
            Tag             =   "_"
            Top             =   2760
            Width           =   5925
            _Version        =   65536
            _ExtentX        =   10451
            _ExtentY        =   556
            Caption         =   "frmManEntidades.frx":1C96
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEntidades.frx":1D02
            Key             =   "frmManEntidades.frx":1D20
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   120
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TrueOleDBList70.TDBCombo tdbcTipoEntidad 
            Height          =   300
            Left            =   2040
            TabIndex        =   24
            Tag             =   "_"
            Top             =   720
            Width           =   2265
            _ExtentX        =   4001
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
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=826"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=762"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1122"
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
            Enabled         =   0   'False
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   299.906
            AutoSize        =   0   'False
            GapHeight       =   36.283
            ListField       =   "Ten_cNombreEntidad"
            BoundColumn     =   "Ten_cTipoEntidad"
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
            _PropDict       =   $"frmManEntidades.frx":1D72
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
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   5895
            TabIndex        =   47
            Top             =   720
            Width           =   1980
            _Version        =   65536
            _ExtentX        =   3492
            _ExtentY        =   556
            Caption         =   "frmManEntidades.frx":1DF9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEntidades.frx":1E65
            Key             =   "frmManEntidades.frx":1E83
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   1
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   5
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TrueOleDBList70.TDBCombo tdbcTipPer 
            Height          =   300
            Left            =   2040
            TabIndex        =   26
            Tag             =   "_"
            Top             =   1230
            Width           =   3465
            _ExtentX        =   6117
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   5144
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=826"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=762"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1122"
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
            Enabled         =   0   'False
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   299.906
            AutoSize        =   0   'False
            GapHeight       =   36.283
            ListField       =   "TipPer"
            BoundColumn     =   "TipPersona"
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
            _PropDict       =   $"frmManEntidades.frx":1ED5
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
         Begin TDBText6Ctl.TDBText tdbtRepresentante 
            Height          =   315
            Left            =   1920
            TabIndex        =   52
            Tag             =   "_"
            Top             =   5640
            Visible         =   0   'False
            Width           =   5895
            _Version        =   65536
            _ExtentX        =   10398
            _ExtentY        =   556
            Caption         =   "frmManEntidades.frx":1F5C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEntidades.frx":1FC8
            Key             =   "frmManEntidades.frx":1FE6
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   80
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin TrueOleDBList70.TDBCombo tdbcEstado 
            Height          =   300
            Left            =   1950
            TabIndex        =   53
            Tag             =   "_"
            Top             =   6150
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2985
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
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=826"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=762"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1122"
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
            Enabled         =   0   'False
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   276.095
            AutoSize        =   -1  'True
            GapHeight       =   36.283
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
            _PropDict       =   $"frmManEntidades.frx":2038
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
         Begin VB.Label Label4 
            Caption         =   "Pais"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   62
            Top             =   4680
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Vinc. Economico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   61
            Top             =   5120
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estado Sunat"
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
            Index           =   13
            Left            =   3660
            TabIndex        =   55
            Top             =   6180
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Representante"
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
            Index           =   9
            Left            =   360
            TabIndex        =   54
            Top             =   5700
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Entidad"
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
            Left            =   495
            TabIndex        =   50
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Persona"
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
            Index           =   12
            Left            =   480
            TabIndex        =   49
            Top             =   1260
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Entidad"
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
            Index           =   8
            Left            =   4455
            TabIndex        =   48
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lbldescripcion 
            Caption         =   "Descripcion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   2790
            Width           =   1095
         End
         Begin VB.Label lblMante 
            AutoSize        =   -1  'True
            Caption         =   "NUEVO REGISTRO"
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
            Left            =   510
            TabIndex        =   9
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":20BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":2499
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":2873
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":2C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":3027
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":3401
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":37DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":3BB5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   960
      Top             =   480
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":4BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":4D29
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":4E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":4FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":5137
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":5291
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":53EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":5545
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":569F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   165
      Top             =   240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":57F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":5D93
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":632D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":68C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":6E61
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":73FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":7995
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":7F2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEntidades.frx":84C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   264
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   3588
      _ExtentX        =   6329
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "imglstTool"
      DisabledImageList=   "imglstdisabled"
      HotImageList    =   "imglstTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Datos F3"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Editar F6"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir F7"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar o Salir ESC"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Public asiento As Boolean       ' *** Para ir al asiento
Public automatico As Boolean    ' *** Para ir a la tabla asiento automatico
Public Cerrar As Boolean
Dim gsGrupo As String
Dim gsRUCAnt As String
Dim Tipodoc As String
Dim lrsDoc As ADODB.Recordset

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub Habilitacontroles(accion As String)

    tdbcTipoEntidad.Enabled = True
    tdbcTipPer.Enabled = True
    tdbcTipoDocumento.Enabled = True
    tdbcEstado.Enabled = True
'    tdbcAplicaconvenio.Enabled = False '--------->NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA

    tdbtNroDocumento.Enabled = True
    tdbtRepresentante.Enabled = True
    tdbtCodigo.Enabled = True
    tdbtDescripcion.Enabled = True
    tdbtDireccion.Enabled = True
'************************************ ------> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
    tdbtApaterno.Enabled = True
    tdbtAmaterno.Enabled = True
    tdbtNombres.Enabled = True
'    OptSi.Enabled = False
'    OptNO.Enabled = False
'************************************ ------> FIN DE REGISTRO
    If accion = "Consulta" Then
    
        Dim sqlPl As String  '************************* NUEVO REGISTRO AÑADIDO 03/07/2013 - PAUL CUEVA
        
        sqlPl = "select * from CNT_Entidad " & _
                "where Emp_cCodigo = '" & gsEmpresa & "' and Ten_cTipoEntidad = '" & tdbcTipoEntidad.BoundText & "'" & _
                " and Ten_cPlame = '1'"
       
         If ExisteDato(sqlPl) = True Then
            fradatos.Visible = True
            fraconvenio.Visible = True
            tdbtDescripcion.Visible = False
            lbldescripcion.Visible = False
'            fradescripcion.Top = 2760
'            Frame2.Height = 4275
'            SSTCentroCosto.Height = 5535
        Else
'            If ExisteDato(sqlPl) = False Then
'            fradatos.Visible = False
'            fraconvenio.Visible = False
'            tdbtDescripcion.Visible = True
'            tdbtDescripcion.Top = 1630
'            lbldescripcion.Top = 1630
'            fradescripcion.Top = 1990
'            Frame2.Height = 3975
'            SSTCentroCosto.Height = 4980
'
'            End If
                       
        End If  '---------------- FIN DEL REGISTRO
    
        tdbcTipoEntidad.Locked = True
        tdbcTipPer.Locked = True
        tdbcTipoDocumento.Locked = True
        tdbcEstado.Locked = True
        tdbcAplicaconvenio.Locked = True

        tdbtNroDocumento.ReadOnly = True
        tdbtRepresentante.ReadOnly = True
        tdbtCodigo.ReadOnly = True
        tdbtDescripcion.ReadOnly = True
        tdbtDireccion.ReadOnly = True

'        tdbtApaterno.ReadOnly = True
'        tdbtAmaterno.ReadOnly = True
'        tdbtNombres.ReadOnly = True
        tdbtNroDocumento.Enabled = False
'        OptSi.Enabled = False
'        OptNO.Enabled = False
'        chkAconvenio.Enabled = False
'        tdbcAplicaconvenio.Locked = False
              
    End If
    
    If accion = "Editar" Then
    
'         Dim sqlPla As String  '************************* NUEVO REGISTRO AÑADIDO 03/07/2013 - PAUL CUEVA
'
'        sqlPla = "select * from CNT_Entidad " & _
'        "where Emp_cCodigo = '" & gsEmpresa & "' and Ten_cTipoEntidad = '" & tdbcTipoEntidad.BoundText & "'" & _
'         " and Ten_cPlame = '1'"
'
'         If ExisteDato(sqlPla) = True Then
'            fradatos.Visible = True
'            fraconvenio.Visible = True
'            tdbtDescripcion.Visible = False
'            lbldescripcion.Visible = False
'            fradescripcion.Top = 2760
'            Frame2.Height = 4275
'            SSTCentroCosto.Height = 5535
'
'        Else
'            If ExisteDato(sqlPla) = False Then
'            fradatos.Visible = False
'            fraconvenio.Visible = False
'            tdbtDescripcion.Visible = True
'            tdbtDescripcion.Top = 1630
'            lbldescripcion.Top = 1630
'            fradescripcion.Top = 1990
'            Frame2.Height = 3975
'            SSTCentroCosto.Height = 4980
'
'            End If
'
'        End If  '---------------- FIN DEL REGISTRO
        
        tdbcTipoEntidad.Locked = True
        tdbcTipPer.Locked = False
        tdbcTipoDocumento.Locked = False
        tdbcEstado.Locked = False
        tdbcAplicaconvenio.Locked = True
        
        tdbtNroDocumento.ReadOnly = False
        tdbtRepresentante.ReadOnly = False
        tdbtCodigo.ReadOnly = True
        tdbtDescripcion.ReadOnly = False
        tdbtDireccion.ReadOnly = False
        
   '********************************************  NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
        tdbtApaterno.ReadOnly = False
        tdbtAmaterno.ReadOnly = False
        tdbtNombres.ReadOnly = False
'        OptSi.Enabled = False
'        OptNO.Enabled = False
'        chkAconvenio.Enabled = False
'        tdbcAplicaconvenio.Locked = False
   
    '****************************************** FIN DE REGISTRO
        tdbtNroDocumento.Enabled = True
    
    End If
    
    If accion = "Nuevo" Then
               
        tdbtApaterno.Text = ""
        tdbtAmaterno.Text = ""
        tdbtNombres.Text = ""
        
        tdbcTipoEntidad.Locked = False
        tdbcTipPer.Locked = False
        tdbcTipoDocumento.Locked = False
        tdbcEstado.Locked = False
        tdbtNroDocumento.ReadOnly = False
        tdbtRepresentante.ReadOnly = False
        tdbtCodigo.ReadOnly = False
        tdbtDescripcion.ReadOnly = False
        tdbtDireccion.ReadOnly = False
    '************************************************ NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
'        tdbtApaterno.Enabled = False
'        tdbtAmaterno.Enabled = False
'        tdbtNombres.Enabled = False
'        OptSi.Enabled = False
'        OptNO.Enabled = False
'        chkAconvenio.Enabled = False
'        tdbcAplicaconvenio.Enabled = False
'        LblApa.Enabled = False
'        lblAma.Enabled = False
'        lblnombres.Enabled = False
    '*********************************************** FIN DE REGISTRO
        tdbtNroDocumento.Enabled = True
      
        
   End If
    
    If accion = "Cancelar" Then
    End If
End Sub

Private Sub chkAconvenio_Click() '----> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA

    If chkAconvenio.Value = "1" Then
       tdbcAplicaconvenio.Enabled = True
    Else
       tdbcAplicaconvenio.Enabled = False
    End If
 
End Sub                      '------------- FIN DE REGISTRO

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSetFocus tdbcTipoEntidad
    If KeyAscii = 48 Then chkActivo.Value = 0
    If KeyAscii = 49 Then chkActivo.Value = 1
End Sub

Private Sub chkTipo_Click()
    
    If chkTipo.Value = vbChecked Then
    
        If CE(tdbcEntidadBus.Text) = "" Then
            Mensajes "Seleccione un tipo de entidad"
            chkTipo.Value = vbUnchecked
            pSetFocus tdbcEntidadBus
        End If
    
        Call CargaTabla(tdbcEntidadBus.BoundText)
    Else
        Call CargaTabla
    End If
      
End Sub

Private Sub chkTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
    If KeyAscii = 48 Then chkTipo.Value = 0
    If KeyAscii = 49 Then chkTipo.Value = 1
End Sub

Private Sub cmdPrint_Click()
    ' *** Abrir el reporte y enviar los parametros
    Dim matriz_fecha(14) As Variant
    cmdPrint.Enabled = False
    
    DoEvents
    matriz_fecha(0) = "@Accion;SEL_PRINT;True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Ent_cCodEntidad;x;True"
    matriz_fecha(3) = "@Ten_cTipoEntidad;x;True"
    matriz_fecha(4) = "@Ent_cPersona;x;True"
    matriz_fecha(5) = "@Ent_cDireccion;x;True"
    matriz_fecha(6) = "@Ent_nRuc;x;True"
    matriz_fecha(7) = "@Ent_cRepresentante;x;True"
    matriz_fecha(8) = "@Ent_cTipoDoc;x;True"
    matriz_fecha(9) = "@Ent_cFlagPersona;x;True"
    matriz_fecha(10) = "@Ent_cEstadoEntidad;x;True"
    matriz_fecha(11) = "@Ent_cEstado;x;True"
    matriz_fecha(12) = "@Ent_cUserCrea;x;True"
    
    matriz_fecha(13) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(14) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    If chkTipo.Value = vbChecked Then
        matriz_fecha(3) = "@Ten_cTipoEntidad;" & tdbcEntidadBus.BoundText & ";True"
    End If
    
    Dim formulas(0) As Variant
    
    If Me.optRuc.Value = True Then
        formulas(0) = "orden = {spCn_GrabaEntidad;1.Ent_nRuc}"
    Else
        If Me.optNombre.Value = True Then
            formulas(0) = "orden = {spCn_GrabaEntidad;1.Ent_cPersona}"
        Else
            formulas(0) = "orden = {spCn_GrabaEntidad;1.Ent_cCodEntidad}"
        End If
    End If
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEntidades.rpt", crptToWindow, "Reporte de Entidades", "", matriz_fecha(), formulas()

    cmdPrint.Enabled = True
End Sub

Private Sub Command1_Click()
Frame3.Visible = False
Frame1.Enabled = True
pSetFocus tdbtCodigoBus
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        Call PosicionFrameReporte
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 200
            .Height = Me.Height - .Top - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 700
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        SSTab1.Width = tdbgCostos.Width
        
        '*** REDIMENSIONAR DETALLE
'        Frame2.Height = Frame1.Height
'        Frame2.Width = Frame1.Width

        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
End Sub

Private Sub ImgCerrar_Click()
Command1_Click
End Sub

Private Sub OptNO_Click()
    If Me.OptNO.Enabled = True Then
        Me.tdbcPais.Enabled = True
        Me.tdbcVinculoEconomico.Enabled = True
    End If
End Sub

Private Sub OptSi_Click()
    If Me.OptSi.Value = True Then
        Me.tdbcPais.Enabled = False
        Me.tdbcVinculoEconomico.Enabled = False
    End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim respuesta As String
    
    Select Case Button.Index
        Case 1:
                DoEvents
                ManNuevo
                
        Case 2: VerDatos
        
        Case 3: Grabar
                If Cerrar = False Then
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
                DoEvents
                
        Case 4: Borrar
                DoEvents
                
        Case 5:
                DoEvents
                Editar
        Case 6: Imprimir
        Case 7
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    DoEvents
                    Call FiltrarRecordSet
                    DoEvents
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                    Form_Resize
                End If
            End If
    End Select
End Sub

Private Sub Borrar()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgCostos.Columns(0).Value) <> "" Then
        ' *** Antes de eliminar verificar q cuenta no haya tenido movimientos
        If VerificaEntidadMvtos = True Then
            Mensajes "Se han registrado movimientos con esta entidad. Elimine movimientos primero...", vbInformation
            Exit Sub
        End If
        ' ***
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            Call CargaArregloMnt
            lArrMnt(0) = "ELIMINAR"                     ' Accion
            lArrMnt(2) = tdbgCostos.Columns(2).Value    ' Codigo de Entidad
            lArrMnt(3) = tdbgCostos.Columns(0).Value    ' Tipo de Entidad
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEntidad", lArrMnt(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Call CargaTabla
            Screen.MousePointer = vbDefault
            DoEvents
            FiltrarRecordSet
            DoEvents
            Mensajes "Registro ha sido eliminado", vbInformation
        End If
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

Private Function VerificaEntidadMvtos() As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    On Error Resume Next
    
    VerificaEntidadMvtos = False
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_GrabaEntidad 'SEL_MVTOS', '" & gsEmpresa & "', '" & tdbgCostos.Columns(2).Value & "', '" & tdbgCostos.Columns(0).Value & "', '', '', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
        If rsArreglo.State = 0 Then
            Mensajes "Seleccione un registro", vbInformation
            Set rsArreglo = Nothing
            Exit Function
        End If
        If rsArreglo(0).Value > 1 Then VerificaEntidadMvtos = True
        Call CerrarRecordSet(rsArreglo)
    End If
    DoEvents
    ' ***
End Function

Private Sub VerDatos()
    On Error Resume Next
    DoEvents
    Call CargaDatosRegistro
    DoEvents
    If lRegElim = False Then
        lblMante = "VER REGISTRO"
        SSTCentroCosto.TabEnabled(1) = True
        SSTCentroCosto.TabEnabled(0) = False
        
        SSTCentroCosto.Tab = 1
        tbrOpciones.Buttons(1).Enabled = False
        tbrOpciones.Buttons(2).Enabled = False
        tbrOpciones.Buttons(4).Enabled = False
        tbrOpciones.Buttons(5).Enabled = False
        
        tbrOpciones.Buttons(7).Image = 8
        chkActivo.Enabled = False
        lTipoMnt = "EDITAR"
        'Call AseguraControl(Me, True)
        Habilitacontroles "Consulta"
        
        'Me.tdbtNroDocumento.Enabled = False
    Else
        lRegElim = False
    End If
End Sub

Private Sub Editar()
    On Error Resume Next
    DoEvents
    GetAllComboNoDomiciliado
    Call CargaDatosRegistro
    DoEvents
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        
'         Dim sqlED As String '--------------------NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
'         sqlED = "select * from CNT_Entidad " & _
'        "where Emp_cCodigo = '" & gsEmpresa & "' and Ten_cTipoEntidad = '" & tdbcTipoEntidad.BoundText & "'" & _
'         " and Ten_cPlame = '1'"
'         If ExisteDato(sqlED) = True Then
'            Me.Height = 6700
'            pSetFocus tdbtDescripcion
'            '-------------------------------
'            fradatos.Visible = True
'            fraconvenio.Visible = True
'            tdbtDescripcion.Visible = False
'            lbldescripcion.Visible = False
'            fradescripcion.Top = 2760
'            Frame2.Height = 4275
'            SSTCentroCosto.Height = 5535
'            '------------------------------
'            tdbcTipoEntidad.Locked = True
'            tdbcTipPer.Locked = False
'            tdbtApaterno.Enabled = True
'            tdbtAmaterno.Enabled = True
'            tdbtNombres.Enabled = True
'            tdbtDescripcion.Enabled = False
'            tdbtDireccion.Enabled = True
'            tdbcTipoDocumento.Enabled = True
'            tdbcTipoDocumento.Locked = False
'            tdbtNroDocumento.Enabled = True
'            OptSi.Enabled = True
'            OptNO.Enabled = True
'            chkAconvenio.Enabled = True
'            tdbcAplicaconvenio.Enabled = True
'        Else
'            Me.Height = 6100
'            pSetFocus tdbtDescripcion
'            '----------------------------------
'            fradatos.Visible = False
'            fraconvenio.Visible = False
'            tdbtDescripcion.Visible = True
'            lbldescripcion.Visible = True
'            tdbtDescripcion.Top = 1630
'            lbldescripcion.Top = 1630
'            fradescripcion.Top = 1990
'            Frame2.Height = 3975
'            SSTCentroCosto.Height = 4980
'            '-----------------------------
'            tdbcTipoEntidad.Locked = False
'            tdbcTipPer.Locked = False
'            tdbtApaterno.Enabled = False
'            tdbtAmaterno.Enabled = False
'            tdbtNombres.Enabled = False
'            tdbtDescripcion.Enabled = True
'            tdbtDireccion.Enabled = True
'            tdbcTipoDocumento.Enabled = False
'            tdbcTipoDocumento.Locked = True
'            tdbtNroDocumento.Enabled = False
'            OptSi.Enabled = False
'            OptNO.Enabled = False
'            chkAconvenio.Enabled = False
'            tdbcAplicaconvenio.Enabled = False
'            End If ' ----------------------------------> FIN DEL REGISTRO
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
         pSetFocus tdbtDescripcion
         tdbcTipoEntidad.Locked = True
         chkActivo.Enabled = True
         Me.tdbtNroDocumento.Enabled = True
         tdbcTipoDocumento.Enabled = True
         tdbcTipoDocumento.Locked = False
    Else
        lRegElim = False
    End If
End Sub

Public Sub ManNuevo()
    
    If gintPercepcion = 1 Then
        Me.chkPorcentajeSunat.Visible = True
    Else
        Me.chkPorcentajeSunat.Visible = False
    End If
        
    On Error Resume Next
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)

    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    DoEvents
    LlenaCombos
    DoEvents
    LlenaAplCOnvenio ' -------------------> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
'    tdbcTipoEntidad_SelChange 0

    chkActivo.Value = 1
    tdbcTipoEntidad.Locked = False
    chkActivo.Enabled = False
    
    Habilitacontroles "Nuevo"
    tdbcTipoEntidad_LostFocus
    pSetFocus tdbcTipoEntidad
    GetAllComboNoDomiciliado
    gsRUCAnt = ""
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    On Error Resume Next
    SSTCentroCosto.TabEnabled(1) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    If Valor = True Then SSTCentroCosto.Tab = 1
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
    tbrOpciones.Buttons(6).Enabled = Not Valor  ' *** imprimir
    If Valor = True Then
        tbrOpciones.Buttons(7).Image = 8
    Else
        tbrOpciones.Buttons(7).Image = 7
    End If
End Sub


Private Sub Cancelar()
    On Error Resume Next
    
    If Me.lblMante = "VER REGISTRO" Then
        Call AseguraControl(Me, False)
    Else
        Call HabilitaControl(Me)
'        Me.Height = 6700
    End If
    Call TabMantenimiento(False)
    pSetFocus tdbgCostos
End Sub

Private Sub PosicionFrameReporte()
    Call Centrar_Objeto(Frame3, Me)
End Sub


Private Sub Imprimir()
    Frame3.Visible = True
    Frame1.Enabled = False
    
    If chkTipo.Value = vbUnchecked Then
        rmBarra.Caption = " Imprimir Entidades "
    Else
        rmBarra.Caption = " Imprimir Entidad : " & tdbcEntidadBus.Text
    End If
    
    pSetFocus Me.cmdPrint
End Sub

Private Sub Grabar()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    
    'If lTipoMnt = "INSERTAR" Then tdbtCodigo = correlativoCodigoEnt(tdbcTipoEntidad.BoundText)
    If validarDatos = False Then Exit Sub
    
    If Not fValidaRUC() Then
       pSetFocus tdbtNroDocumento
       Exit Sub
    End If
    
    If Me.OptNO.Value = True Then
        If Me.tdbcPais.BoundText = vbNullString And gstrVersionLE = "1" Then
            MsgBox "Debe seleccionar un Pais", vbExclamation, "Mensaje"
            pSetFocus tdbtNroDocumento
            Exit Sub
       
        End If
    End If
    
    If Me.OptNO.Value = True Then
        If Me.tdbcVinculoEconomico.BoundText = vbNullString And gstrVersionLE = "1" Then
            MsgBox "Debe seleccionar un Vinculo Economico", vbExclamation, "Mensaje"
            pSetFocus tdbcVinculoEconomico
        End If
    End If
    
    If tdbcAplicaconvenio.Enabled = True Then
        If Me.tdbcAplicaconvenio.BoundText = vbNullString And gstrVersionLE = "1" Then
            MsgBox "Debe seleccionar el Pais de Convenio", vbExclamation, "Mensaje"
            pSetFocus tdbcAplicaconvenio
            Exit Sub
        End If
    End If
    
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEntidad", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    On Error Resume Next
    
    Call Cancelar
    
    CargaTabla
    ' *** Buscar el codigo creado y posicionarse alli
    Dim Valor As Integer
    Valor = BuscarEntRs(tdbcTipoEntidad.BoundText, tdbtCodigo, lrsTabla, 2, 1)
    If Valor = 0 Then lrsTabla.MoveFirst
    ' ***
    FiltrarRecordSet
    Mensajes "Los datos se grabaron con exito...", vbInformation
    DoEvents
    pSetFocus tdbgCostos
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    Cerrar = False
    gsCodigoEnt = ""
    If asiento = True Then
        On Error Resume Next
        frmManAsientosContables.Enabled = True
        frmManAsientosContables.tdbgDetalle.Columns(22) = Me.tdbcTipoEntidad.BoundText
        frmManAsientosContables.tdbgDetalle.Columns(4) = Me.tdbtCodigo
        frmManAsientosContables.tdbgDetalle.Columns(23) = Me.tdbtDescripcion
        frmManAsientosContables.tdbgDetalle.Update
        gsCodigoEnt = Me.tdbtCodigo
        Cerrar = True
        DoEvents
        Unload Me
        DoEvents
        
        frmManAsientosContables.tdbgDetalle.Col = frmManAsientosContables.BuscaCeldaActiva(4)
        pSetFocus frmManAsientosContables.tdbgDetalle
        DoEvents
        Cerrar = True
        
        On Error GoTo 0
    End If
    
    If automatico = True Then
        On Error Resume Next
        frmBusTipoAsiento.Enabled = True
        frmBusTipoAsiento.tdbcTipoEntidad.BoundText = Me.tdbcTipoEntidad.BoundText
        frmBusTipoAsiento.tdbtEntidad = Me.tdbtCodigo
        frmBusTipoAsiento.tdbtNombreEntidad = Me.tdbtDescripcion
        gsCodigoEnt = Me.tdbtCodigo
        Unload Me
        DoEvents
        pSetFocus frmBusTipoAsiento.tdbtEntidad
        pSendKeys "{Enter}"
        Cerrar = True
        On Error GoTo 0
    End If
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function BuscarEntRs(Tipo As String, cadenas As String, rs As Recordset, colTipo As Integer, Col As Integer) As Integer
    
    On Error Resume Next
    
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Dim Fila As Integer
            BuscarEntRs = 0
            Fila = 0
            rs.MoveFirst
            Do While Not rs.EOF
                If Trim(Tipo) = Trim(rs(colTipo).Value) And Trim(cadenas) = Trim(rs(Col).Value) Then
                    BuscarEntRs = Fila
                    Exit Function
                End If
                rs.MoveNext
                Fila = Fila + 1
            Loop
        End If
    End If
End Function

Private Function ValidaRUCJuridicNat() As Boolean
    ValidaRUCJuridicNat = True
'    If tdbtNroDocumento.Enabled Then
'
'        If tdbcTipPer.BoundText = "J" And _
'           tdbcTipoDocumento.BoundText = "04" And _
'           CE(tdbtNroDocumento) <> "" And _
'           Left(CE(tdbtNroDocumento.Text), 1) <> "2" Then
'
'            Mensajes "El RUC no es valido para una persona de tipo Juridica", vbInformation
'            pSetFocus tdbtNroDocumento
'            ValidaRUCJuridicNat = False
'        End If
'
'        If tdbcTipPer.BoundText = "N" And _
'           tdbcTipoDocumento.BoundText = "04" And _
'           CE(tdbtNroDocumento) <> "" And _
'           Left(CE(tdbtNroDocumento.Text), 1) <> "1" Then
'
'            Mensajes "El RUC no es valido para una persona de tipo Natural", vbInformation
'            pSetFocus tdbtNroDocumento
'            ValidaRUCJuridicNat = False
'        End If
'    End If
End Function

Private Function validarDatos() As Boolean
    validarDatos = False

    If TextoLleno2(tdbtCodigo, "Codigo") = False Then Exit Function
    
    If tdbtDescripcion.Enabled = True And tdbcTipPer.BoundText <> "N" Then  '-----------> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
        If TextoLleno2(tdbtDescripcion, "Descripción") = False Then Exit Function
    Else
        
        If TextoLleno2(tdbtApaterno, "Apellido Paterno") = False Then Exit Function
        If tdbcTipPer.BoundText <> "X" Then
        If TextoLleno2(tdbtAmaterno, "Apellido Materno") = False Then Exit Function
        
        End If
        If TextoLleno2(tdbtNombres, "Nombres") = False Then Exit Function
     
        
    End If ' ------------FIN DEL REGISTRO
    
    If gsRUCAnt <> CE(tdbtNroDocumento.Text) Then
        If VerificaRuc(Me.tdbcTipoEntidad.BoundText) <> "" Then
            Mensajes "Ya existe este número de, " & UCase(CE(tdbcTipoDocumento.Text)), vbInformation
            Exit Function
        End If
    End If
    
    If ValidaRUCJuridicNat = False Then
        Exit Function
    End If
        
    Dim Digitos  As Integer
    Digitos = BuscaTamanioDoc(tdbcTipoDocumento.BoundText)
    If Len(CE(tdbtNroDocumento.Text)) <> Digitos Then
        Mensajes "La cantidad de digitos del " & tdbcTipoDocumento.Text & " no es valido, el número permitido es " & Digitos, vbInformation
        pSetFocus tdbtNroDocumento
        Exit Function
    End If
    
    If Digitos > 0 Then
        If TextoLleno2(tdbtNroDocumento, "Documento") = False Then Exit Function
    End If
    
    Dim sqlSp   As String
    Dim clDatos As clsMantoTablas
    Dim rsAux   As New ADODB.Recordset
    Dim arrDatos() As Variant

    On Local Error GoTo ErrorEjecucion
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCNT_ENTIDAD_DOCU 'BUSCARREGISTRO', '" & gsEmpresa & "', '" & tdbcTipoEntidad.Columns(0).Value & "','" & tdbcTipPer.Columns(0).Value & "', '" & tdbcTipoDocumento.Columns(0).Value & "'"
    arrDatos = Array(sqlSp)
    Set rsAux = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    'If Not rsAux Is Nothing Then
    If rsAux.State <> adStateClosed Then
        If Not (rsAux.EOF And rsAux.BOF) Then
            If Not rsAux.State = 0 Then
                ' *** Llenar grilla con el RecordSet
               If Len(Trim(rsAux!Ten_cTipoEntidad)) > 0 Then
                  If Len(Trim(tdbcTipoEntidad.BoundText)) = 0 Then
                     Mensajes "Seleccionar Entidad", vbInformation
                     pSetFocus tdbcTipoEntidad
                     Exit Function
                  End If
               End If
               If Len(Trim(rsAux!Edoc_cTipoPersona)) > 0 Then
                  If Len(Trim(tdbcTipPer.BoundText)) = 0 Then
                     Mensajes "Seleccionar Tipo Persona", vbInformation
                     pSetFocus tdbcTipPer
                     Exit Function
                  End If
               End If
               If Len(Trim(rsAux!Edoc_cTipoDoc)) > 0 Then
                  If Len(Trim(tdbcTipoDocumento.BoundText)) = 0 Then
                     Mensajes "Seleccionar Tipo Documento", vbInformation
                     pSetFocus tdbcTipoDocumento
                     Exit Function
                  Else
                     If Len(Trim(tdbtNroDocumento.Text)) = 0 Then
                        Mensajes "Ingrese Nro. de Documento", vbInformation
                        pSetFocus tdbtNroDocumento
                        Exit Function
                     Else
                        If Trim(tdbcTipoDocumento.Text) = "RUC" Then
                           If Not fValidarNroRuc(tdbtNroDocumento.Text) Then
                              Mensajes "El RUC no es valido", vbInformation
                              pSetFocus tdbtNroDocumento
                              Exit Function
                           End If
                        End If
                     End If
                  End If
               End If
            Else
                Mensajes "Configurar Entidad - Documento, para ejecutar esta opción", vbOKOnly + vbInformation
                Exit Function
            End If
            CerrarRecordSet rsAux
            
        End If
    End If
    validarDatos = True
    
    Exit Function
    
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Function
    
Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    'ReDim lArrMnt(21) As Variant   '-------> modificado el numero 12 por 18 NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
    ReDim lArrMnt(22) As Variant
    lArrMnt(0) = lTipoMnt           ' Accion
    lArrMnt(1) = gsEmpresa          ' Empresa
    lArrMnt(2) = tdbtCodigo         ' Codigo
    lArrMnt(3) = tdbcTipoEntidad.BoundText    ' Codigo
    lArrMnt(4) = IIf(tdbtDescripcion = vbNullString, tdbtApaterno & " " & tdbtAmaterno & " " & tdbtNombres, tdbtDescripcion)    ' Nombre o Razon Social
    lArrMnt(5) = tdbtDireccion      ' Direccion
    lArrMnt(6) = tdbtNroDocumento   ' Numero de Documento
    lArrMnt(7) = tdbtRepresentante  ' Representante
    lArrMnt(8) = Trim(tdbcTipoDocumento.BoundText)      ' Tipo de Documento
    ' *** Tipo de Persona
    lArrMnt(9) = tdbcTipPer.BoundText
    ' *** Estado Sunat
    lArrMnt(10) = Trim(tdbcEstado.BoundText)
    ' *** Estado de Entidad
    If chkActivo.Value = 1 Then
        lArrMnt(11) = "A"
    Else
        lArrMnt(11) = "I"
    End If
    lArrMnt(12) = gsUsuario         ' Usuario
    
    lArrMnt(13) = tdbtApaterno '----------------------------> NUEVOS REGISTROS AÑADIDO 02/07/2013 - PAUL CUEVA
    lArrMnt(14) = tdbtAmaterno
    lArrMnt(15) = tdbtNombres
    
    If OptSi.Value = True Then
       lArrMnt(16) = 1
    Else
       lArrMnt(16) = 2
    End If
    
    If chkAconvenio.Value = "1" Then
        lArrMnt(17) = Trim(tdbcAplicaconvenio.BoundText)
    Else
        lArrMnt(17) = ""
    End If '-------------------------> FIN DEL REGISTRO
    
    lArrMnt(18) = ""
    lArrMnt(19) = IIf(Me.tdbcPais.BoundText = vbNullString, Null, Me.tdbcPais.BoundText)
    lArrMnt(20) = IIf(Me.tdbcVinculoEconomico.BoundText = vbNullString, Null, Me.tdbcVinculoEconomico.BoundText)
    lArrMnt(21) = IIf(Me.tdbcAplicaconvenio.BoundText = vbNullString, Null, Me.tdbcAplicaconvenio.BoundText)
    lArrMnt(22) = Me.chkPorcentajeSunat.Value
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar
        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub
Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    On Error Resume Next
    Centrar_form Me
    ' *** Llenando las grillas y los combos
    LlenaCombos
    LlenarTipPer
    LlenarTipDocu
    CargaTabla
    LlenaAplCOnvenio '---------------> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
    ' *** Seteando algunas variables
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    SSTCentroCosto.Tab = 0
    asiento = False
    automatico = False
    
    If chkAconvenio.Value = 0 Then Me.tdbcAplicaconvenio.Enabled = False
    GetAllComboNoDomiciliado
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
End Sub
Private Sub LlenaCombos()
    On Error Resume Next
    Dim sqlcombos As String
    ' *** Llenando el tipo de Entidad
    sqlcombos = "SELECT Ten_cTipoEntidad, Ten_cNombreEntidad From CNT_ENTIDAD " & _
                " WHERE Emp_cCodigo = '" & gsEmpresa & "' ORDER BY Ten_cNombreEntidad"
    LlenarComboAddItem tdbcTipoEntidad, sqlcombos
    LlenarComboAddItem tdbcEntidadBus, sqlcombos
    
    ' *** Llenando el Estado de Entidades
    sqlcombos = "SELECT RTRIM(Tab_cCodigo), Tab_cDescripCampo FROM TABLA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Tab_cEstado = 'A' " & _
                "AND Tab_cTabla = '018' ORDER BY Tab_cCodigo"
    LlenarComboAddItem tdbcEstado, sqlcombos
End Sub

Public Sub LlenaCombosTipoPerDocu()
    On Error Resume Next
    LlenarTipPer
    DoEvents
    LlenarTipDocu
    DoEvents
    tdbcTipPer.ReBind
    tdbcTipoDocumento.ReBind
    DoEvents
End Sub

Private Sub LlenarTipPer()
    Dim sqlcombos As String
    On Error Resume Next
    ' *** Llenando el Estado de Entidades
    sqlcombos = "SELECT Rtrim(Isnull(A.Edoc_cTipoPersona,'')) TipPersona,C.Tab_cDescripCampo AS TipPer " & _
                "FROM CNT_ENTIDAD_DOCU A " & _
                "LEFT JOIN CNT_ENTIDAD B ON A.Ten_cTipoEntidad=B.Ten_cTipoEntidad And B.Emp_cCodigo = '" & gsEmpresa & "' " & _
                "LEFT JOIN TABLA C ON A.Edoc_cTipoPersona=C.Tab_cCodigo And C.Emp_cCodigo = '" & gsEmpresa & "' And C.Tab_cTabla = '039' " & _
                "LEFT JOIN TABLA D ON A.Edoc_cTipoDoc=D.Tab_cCodigo And D.Emp_cCodigo = '" & gsEmpresa & "' And D.Tab_cTabla = '003' " & _
                "WHERE A.Emp_cCodigo = '" & gsEmpresa & "' And A.Edoc_cDeleted<>'*' And A.Ten_cTipoEntidad='" & tdbcTipoEntidad.Columns(0).Value & "' " & _
                "GROUP BY A.Edoc_cTipoPersona,C.Tab_cDescripCampo " & _
                "Order by A.Edoc_cTipoPersona"
    LlenarComboAddItem tdbcTipPer, sqlcombos
    tdbcTipPer.ReBind
End Sub
Private Sub LlenarTipDocu()
    Dim sqlcombos As String
    Dim nReg As Integer
    On Error Resume Next
    Set tdbcTipoDocumento.RowSource = Nothing
    CerrarRecordSet lrsDoc
    
    sqlcombos = "SELECT A.Edoc_cTipoDoc,D.Tab_cDescripCampo AS TipDoc,Convert(Char(2),D.Tab_nLongitud) As Tamanio " & _
                "FROM CNT_ENTIDAD_DOCU A " & _
                "LEFT JOIN CNT_ENTIDAD B ON A.Ten_cTipoEntidad=B.Ten_cTipoEntidad And B.Emp_cCodigo = '" & gsEmpresa & "' " & _
                "LEFT JOIN TABLA C ON A.Edoc_cTipoPersona=C.Tab_cCodigo And C.Emp_cCodigo = '" & gsEmpresa & "' And C.Tab_cTabla = '039' " & _
                "LEFT JOIN TABLA D ON A.Edoc_cTipoDoc=D.Tab_cCodigo And D.Emp_cCodigo = '" & gsEmpresa & "' And D.Tab_cTabla = '003' " & _
                "WHERE A.Emp_cCodigo = '" & gsEmpresa & "' And A.Edoc_cDeleted<>'*' And A.Ten_cTipoEntidad='" & tdbcTipoEntidad.Columns(0).Value & "' And A.Edoc_cTipoPersona='" & tdbcTipPer.Columns(0).Value & "' " & _
                " Order by D.Tab_cDescripCampo desc"
    
    'LlenarComboAddItem tdbcTipoDocumento, sqlcombos
    Call LlenarRecordSet(sqlcombos, lrsDoc)

    If Not lrsDoc Is Nothing Then
        nReg = lrsDoc.RecordCount
        If nReg > 10 Then nReg = 10
        tdbcTipoDocumento.BoundColumn = "Edoc_cTipoDoc"
        tdbcTipoDocumento.ListField = "TipDoc"
        tdbcTipoDocumento.MaxComboItems = nReg
        tdbcTipoDocumento.DropDownWidth = 4500
        Set tdbcTipoDocumento.RowSource = lrsDoc
        
        'tdbcTipoDocumento.Bookmark = 1
        
    Else
        tdbcTipoDocumento.BoundText = ""
    End If
    DoEvents
    'tdbcTipoDocumento.Columns(0).DataField = "Edoc_cTipoDoc"
    'tdbcTipoDocumento.Columns(1).DataField = "TipDoc"
    'tdbcTipoDocumento.Columns(3).DataField = "Tamanio"
    'tdbcTipoDocumento.BoundColumn = "Edoc_cTipoDoc"
    'tdbcTipoDocumento.ListField = "TipDoc"

    'Set Me.tdbcTipoDocumento.DataSource = lrsDoc
    
    'tdbdOperaTC.Columns(1).Width = 0
    'tdbdOperaTC.Columns(1).Visible = False
        
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Call CerrarRecordSet(lrsTabla)
    
    If asiento = True Then
        frmManAsientosContables.Enabled = True
    End If
    
    If automatico = True Then
        frmBusTipoAsiento.Enabled = True
    End If

    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub
Private Sub CargaTabla(Optional cTipoentidad As String = "")
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
        
    On Error Resume Next
        
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaEntidad 'SEL_ALL', '" & gsEmpresa & "', '', '" & cTipoentidad & "', '', '', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then
        If Not (lrsTabla.EOF And lrsTabla.BOF) Then
            lrsTabla.Sort = "Ten_cTipoEntidad, Ent_cCodEntidad"
            tdbgCostos.DataSource = lrsTabla
        End If
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    
    On Local Error GoTo ErrorEjecucion
    
    sqlSp = "spCn_GrabaEntidad 'SEL_REG', '" & gsEmpresa & "', '" & tdbgCostos.Columns(2).Value & "', '" & tdbgCostos.Columns(0).Value & "', '', '', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro el registro. Probablemente eliminado desde otra sesion", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    
    ' *** Asignando Datos de la Entidad
    tdbtCodigo = CE(rsArreglo!Ent_cCodEntidad)
    tdbcTipoEntidad.BoundText = CE(rsArreglo!Ten_cTipoEntidad)
    DoEvents
    LlenarTipPer
    DoEvents
    tdbtDescripcion = CE(rsArreglo!Ent_cPersona)
    tdbtDireccion = CE(rsArreglo!Ent_cDireccion)
    tdbcTipPer.BoundText = CE(rsArreglo!Ent_cFlagPersona)
    tdbtApaterno = CE(rsArreglo!Ent_capaterno) '-----> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
    tdbtAmaterno = CE(rsArreglo!Ent_camaterno)
    tdbtNombres = CE(rsArreglo!Ent_cnombres) '------> FIN DEL REGISTRO
    DoEvents
    LlenarTipDocu
    DoEvents
    
    Tipodoc = CE(rsArreglo!Ent_cTipoDoc)
    tdbcTipoDocumento.BoundText = CE(rsArreglo!Ent_cTipoDoc)
    tdbtNroDocumento = CE(rsArreglo!Ent_nRuc)
    tdbtRepresentante = CE(rsArreglo!Ent_cRepresentante)
    tdbcEstado.BoundText = CE(rsArreglo!Ent_cEstadoEntidad)
    gsRUCAnt = CE(rsArreglo!Ent_nRuc)
    '--------------------------> NUEVO REGISTRO AÑADIDO
    If CE(rsArreglo!ent_cFlagDomiciliado) = "1" Then
        OptSi.Value = True
    Else
        OptNO.Value = True
    End If
    
'    If CE(rsArreglo!Ent_cAconvenio) <> "" Then
'        chkAconvenio.Value = "1"
'        tdbcAplicaconvenio.BoundText = CE(rsArreglo!Ent_cAconvenio)
'        tdbcAplicaconvenio.Enabled = True
'    Else
'        chkAconvenio.Value = "0"
'    End If

    
    '--------------------------> FIN DEL REGISTRO
    If CE(rsArreglo!Ent_cEstado) = "A" Then
        chkActivo.Value = 1
    Else
        chkActivo.Value = 0
    End If
   
    Me.tdbcPais.BoundText = CE(rsArreglo!Id_Pais)
    Me.tdbcVinculoEconomico.BoundText = CE(rsArreglo!Id_Vinculo_Economico)
    Me.tdbcAplicaconvenio.BoundText = CE(rsArreglo!Id_Convenio)
    Me.chkPorcentajeSunat.Value = NE(rsArreglo!PorcentajeSunat)
    Call CerrarRecordSet(rsArreglo)
    ' ***
    
    Exit Sub
    
ErrorEjecucion:
    'Mensajes Str(Err.Number) & Err.Description, vbInformation
    
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    tdbcEntidadBus.ReBind
    
    DoEvents
    Dim cadena As String
    Dim filtros(4) As String
    Dim i As Integer
    
    On Error GoTo serror
    If lrsTabla Is Nothing Then Exit Sub
    cadena = ""
    tdbcEntidadBus.ReBind
    DoEvents
    If CE(tdbtCodigoBus) <> "" Then filtros(0) = "Ent_cCodEntidad like '*" & tdbtCodigoBus & "*'"
    If CE(tdbtDescripcionBus) <> "" Then filtros(1) = "Ent_cPersona like '*" & tdbtDescripcionBus & "*'"
    If CE(tdbtNroDocBus) <> "" Then filtros(2) = "Ent_nRuc like '*" & tdbtNroDocBus & "*'"
    If CE(tdbtDireccionBus) <> "" Then filtros(3) = "Ent_cDireccion like '*" & tdbtDireccionBus & "*'"
    If chkTipo.Value = 1 Then filtros(4) = "Ten_cTipoEntidad like '*" & tdbcEntidadBus.BoundText & "*'"
    For i = 0 To 4
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    lrsTabla.Filter = 0
    ' *** Filtrando segun campos
    If Not lrsTabla Is Nothing Then
        If Not (lrsTabla.EOF And lrsTabla.BOF) Then
            If Trim(cadena) <> "" Then
                lrsTabla.Filter = cadena
            Else
                lrsTabla.Filter = 0
            End If
        End If
    End If
    
    If cadena = "" Then lrsTabla.Filter = 0
    Exit Sub
serror:
    
End Sub

Private Sub tdbcEntidadBus_ItemChange()
    If chkTipo.Value = vbChecked Then
        
        Call CargaTabla(tdbcEntidadBus.BoundText)
        
    End If
End Sub

Private Sub tdbcEntidadBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtDescripcionBus
End If
End Sub

Private Sub tdbcEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbcTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub
Private Sub tdbcTipoEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub
Private Sub tdbcTipoEntidad_LostFocus()

    ' *** Hallar el correlativo de la Entidad
    If lTipoMnt = "INSERTAR" Then tdbtCodigo = correlativoCodigoEnt(tdbcTipoEntidad.BoundText)
    If lblMante <> "VER REGISTRO" Then
        
        LlenarTipPer
        DoEvents
        LlenarTipDocu
        DoEvents
    End If
    
End Sub
Private Sub tdbcTipoEntidad_SelChange(Cancel As Integer)
On Error Resume Next
    tdbcTipPer.Text = ""
    tdbcTipoDocumento.Text = ""
    tdbtNroDocumento.Text = ""
    'LlenarTipPer
    'LlenarTipDocu
      
     Dim sqlPl As String  '************************* NUEVO REGISTRO AÑADIDO 03/07/2013 - PAUL CUEVA
        
        sqlPl = "select * from CNT_Entidad " & _
        "where Emp_cCodigo = '" & gsEmpresa & "' and Ten_cTipoEntidad = '" & tdbcTipoEntidad.BoundText & "'" & _
         " and Ten_cPlame = '1'"
       
         If ExisteDato(sqlPl) = True Then

            fradatos.Visible = True
            fraconvenio.Visible = True
            
            tdbtDescripcion.Visible = False
            lbldescripcion.Visible = False
            Me.Height = 6700
            fradescripcion.Top = 2760
            Frame2.Height = 4275
            SSTCentroCosto.Height = 5535
                                                           
            tdbtApaterno.Enabled = True
            tdbtAmaterno.Enabled = True
            tdbtNombres.Enabled = True
            'chkDomiciliado.Enabled = True
            OptSi.Enabled = True
            OptNO.Enabled = True
            chkAconvenio.Enabled = True
'            tdbcAplicaconvenio.Enabled = False     '-------------> Nuevo habilitando
            
'            If chkAconvenio.Value = 1 Then
'            tdbcAplicaconvenio.Locked = False
'            Else
'            tdbcAplicaconvenio.Locked = True
'            End If

            LblApa.Enabled = True
            lblAma.Enabled = True
            lblnombres.Enabled = True
            tdbtDescripcion.Enabled = False
            lbldescripcion.Enabled = False
            
        Else
            If ExisteDato(sqlPl) = False Then
            
'            fradatos.Visible = False
'            fraconvenio.Visible = False
            tdbtDescripcion.Visible = True
            lbldescripcion.Visible = True
                        
'            tdbtDescripcion.Top = 1630
'            lbldescripcion.Top = 1630
'            fradescripcion.Top = 1990
'            Frame2.Height = 3975
'            Me.Height = 6100
'            SSTCentroCosto.Height = 4980
                        
'            Frame2.Top = 4500
                   
            tdbtApaterno.Enabled = False
            tdbtAmaterno.Enabled = False
            tdbtNombres.Enabled = False
            'chkDomiciliado.Enabled = False
'            OptSi.Enabled = False
'            OptNO.Enabled = False
'            chkAconvenio.Enabled = False
'            tdbcAplicaconvenio.Enabled = False      '-----------> nuevo habilitando
'            LblApa.Enabled = False
'            lblAma.Enabled = False
'            lblnombres.Enabled = False
            tdbtDescripcion.ReadOnly = False
            tdbtDescripcion.Enabled = True
            lbldescripcion.Enabled = True
                        
            End If
                       
        End If  '---------------- FIN DEL REGISTRO
   
End Sub

Private Sub tdbcTipPer_ItemChange()
    If lblMante <> "VER REGISTRO" Then
        LlenarTipDocu
    End If
    
    OptSi.Value = True
    
    If tdbcTipPer.BoundText = "N" Then
'        OptSi.Value = True
        Me.tdbtAmaterno.Enabled = True
        Me.tdbtApaterno.Enabled = True
        Me.tdbtNombres.Enabled = True
    Else
'        OptNO.Value = True
        Me.tdbtAmaterno.Enabled = False
        Me.tdbtApaterno.Enabled = False
        Me.tdbtNombres.Enabled = False
    End If
End Sub
Private Sub tdbcTipPer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbgCostos_GotFocus()
tdbgCostos.HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgCostos_HeadClick(ByVal ColIndex As Integer)
If Not lrsTabla Is Nothing Then
    If Not (lrsTabla.EOF And lrsTabla.BOF) Then
        If lrsTabla.RecordCount > 0 Then
        
            lrsTabla.Sort = tdbgCostos.Columns(ColIndex).DataField
            tdbgCostos.DataSource = lrsTabla
            
        End If
    End If
End If
End Sub

Private Sub tdbgCostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Editar
End If
End Sub

Private Sub tdbgCostos_LostFocus()
tdbgCostos.HighlightRowStyle = ""
End Sub

Private Sub tdbtAmaterno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbtApaterno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbtCodigoBus_Change()
    Call FiltrarRecordSet
End Sub
Private Sub tdbtCodigo_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbtCodigo.ReadOnly = True
    Else
        tdbtCodigo.ReadOnly = False
    End If
    ' ***
End Sub

Private Sub tdbtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDescripcion = Replace(tdbtDescripcion, "'", "")
       tdbtDescripcion.SelStart = Len(tdbtDescripcion)
    End If
End Sub

Private Sub tdbtDescripcionBus_Change()
    
    If gsKey = 219 Then
       tdbtDescripcionBus = Replace(tdbtDescripcionBus, "'", "")
       tdbtDescripcionBus.SelStart = Len(tdbtDescripcionBus)
    End If
    
    Call FiltrarRecordSet
End Sub

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

Private Sub tdbtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDireccion = Replace(tdbtDireccion, "'", "")
       tdbtDireccion.SelStart = Len(tdbtDireccion)
    End If
End Sub

Private Sub tdbtDireccionBus_Change()
    Call FiltrarRecordSet
End Sub

Private Sub tdbtDireccionBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbgCostos
End If
End Sub

Private Sub tdbtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub tdbtNroDocBus_Change()
    Call FiltrarRecordSet
End Sub
Private Sub tdbtNroDocumento_GotFocus()
    Me.tdbtNroDocumento.MaxLength = Val(IIf(IsNull(Me.tdbcTipoDocumento.Columns(2).Value), "0", Me.tdbcTipoDocumento.Columns(2).Value))
    If tdbtNroDocumento.MaxLength = 0 Then
       tdbtNroDocumento.ReadOnly = True
    Else
       tdbtNroDocumento.ReadOnly = False
    End If
End Sub
Private Sub tdbtNroDocumento_LostFocus()
    Dim varCodigo As String
    
    If tdbtNroDocumento.MaxLength = 0 Then
       tdbtNroDocumento = ""
       Exit Sub
    End If
    
    If tbrOpciones.Buttons(3).Enabled = True Then
        If Not fValidaRUC() Then tdbtNroDocumento.SetFocus: Exit Sub
    End If
    
    If Me.tdbtNroDocumento.Enabled = True And Trim(tdbtNroDocumento) <> "" Then
       If Len(CE(tdbtNroDocumento)) <> NE(tdbcTipoDocumento.Columns(2).Value) Then
            Mensajes "Tipo de Documento no coincide con la cantidad de digitos. Digito permitido " & CE(tdbcTipoDocumento.Columns(2).Value), vbInformation
            pSetFocus tdbtNroDocumento
            Exit Sub
        End If
    End If
    
    If tdbcTipoDocumento.BoundText = "04" And tdbtNroDocumento.Enabled = True Then
        ' Verifica si es correcto el nro de RUC
        If Not fValidarNroRuc(tdbtNroDocumento) And CE(tdbtNroDocumento.Text) <> "" Then
           MsgBox "El RUC no es valido", vbInformation
           pSetFocus tdbtNroDocumento
           Exit Sub
        End If
        
        ' *** Verificar q ruc no exista
        If CE(Me.tdbtNroDocumento.Text) <> "" Then
            varCodigo = VerificaRuc(Me.tdbcTipoEntidad.BoundText)
        Else
            varCodigo = ""
        End If
                
        If lTipoMnt = "INSERTAR" Then
            If varCodigo <> "" Then
                Mensajes "Nro de Ruc ya existe en el codigo: " & varCodigo & " . Verificar", vbInformation
                pSetFocus tdbtNroDocumento
            End If
        Else
            If varCodigo <> "" And varCodigo <> Me.tdbtCodigo Then
                Mensajes "Nro de Ruc ya existe en el codigo: " & varCodigo & " . Verificar", vbInformation
                pSetFocus tdbtNroDocumento
            End If
        End If
        ' ***
    End If
End Sub

Private Function fValidaRUC() As Boolean
    fValidaRUC = False
    
    If Trim(tdbtNroDocumento) <> "" And tdbcTipoDocumento.BoundText = "04" Then
        If Len(Trim(tdbtNroDocumento)) <> 11 And Me.tdbtNroDocumento.Enabled = True Then
            Mensajes "Numero de digitos de Ruc debe ser igual a 11. Verificar.. ", vbInformation
            Exit Function
        End If
        
        ' Verifica si es correcto el nro de RUC
        If Not fValidarNroRuc(tdbtNroDocumento) And tdbcTipoDocumento.BoundText = 4 Then
           MsgBox "El RUC no es válido", vbInformation
           Exit Function
        End If
    End If
    
    fValidaRUC = True
End Function

Private Function VerificaRuc(Tipo As String) As String

    Dim rsCosto As New ADODB.Recordset
    Dim sqlver As String
    
    VerificaRuc = ""
    If Trim(Me.tdbcTipoDocumento.BoundText) <> "00" Then
        sqlver = "select Ent_cCodEntidad from dbo.CNM_ENTIDAD WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
                 "AND Ten_cTipoEntidad =  '" & Tipo & "' " & _
                 "AND Ent_nRuc = '" & Me.tdbtNroDocumento & "' "
'    Else
'        sqlver = "select Ent_cCodEntidad from dbo.CNM_ENTIDAD WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
'                 "AND Ten_cTipoEntidad =  '" & Tipo & "' " & _
'                 "AND Ent_nRuc = '" & Me.tdbtNroDocumento & "' " & _
'                 "and Ent_cTipoDoc = '" & Trim(Me.tdbcTipoDocumento.BoundText) & "' and Ent_cFlagPersona = '" & Trim(Me.tdbcTipPer.BoundText) & "'"
    Call Conectar
    rsCosto.Open sqlver, gcnSistema
    If Not rsCosto.EOF And Not rsCosto.BOF Then VerificaRuc = rsCosto(0).Value
    Call CerrarRecordSet(rsCosto)
    Call Desconectar
    End If

End Function

Private Sub tdbtRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtRepresentante = Replace(tdbtRepresentante, "'", "")
       tdbtRepresentante.SelStart = Len(tdbtRepresentante)
    End If
End Sub

Private Sub GetAllComboNoDomiciliado()
    Dim strSql As String
    
    strSql = "SELECT p.Id_Pais, p.Nom_Pais FROM dbo.Pais P"
    LlenarComboAddItem Me.tdbcPais, strSql
    
    strSql = "SELECT ve.Id_Vinculo_Economico, ve.Descrip_Vinculo_Economico FROM dbo.Vinculo_Economico VE"
    LlenarComboAddItem Me.tdbcVinculoEconomico, strSql
    
End Sub

Private Sub LlenaAplCOnvenio() '-----------> NUEVO REGISTRO AÑADIDO 02/07/2013 - PAUL CUEVA
    On Error Resume Next
    Dim sqlcombos As String
   
'sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo " & _
'      "From TABLA " & _
'          "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND " & _
'          "Tab_cTabla='094' ORDER BY Tab_cDescripCampo "
    sqlcombos = "SELECT c.Id_Convenio, c.Nom_Convenio FROM dbo.Convenio C"

    LlenarComboAddItem tdbcAplicaconvenio, sqlcombos
    tdbcAplicaconvenio.ReBind
    
End Sub  '--------------------------------->  FIN DEL REGISTRO

