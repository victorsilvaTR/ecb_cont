VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManIndicadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Indicadores"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   Icon            =   "frmManIndicadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   10500
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   4890
      Left            =   45
      TabIndex        =   21
      Top             =   405
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   8625
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consultas de Indicadores"
      TabPicture(0)   =   "frmManIndicadores.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Indicadores"
      TabPicture(1)   =   "frmManIndicadores.frx":0EE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4410
         Left            =   90
         TabIndex        =   25
         Top             =   405
         Width           =   10215
         Begin VB.Frame Frame3 
            Height          =   2940
            Left            =   180
            TabIndex        =   31
            Top             =   1350
            Width           =   9945
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
               Height          =   735
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               Top             =   2070
               Width           =   4470
            End
            Begin VB.ComboBox tdbcOperador 
               Height          =   315
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   1440
               Width           =   2880
            End
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
               Left            =   1170
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   240
               Width           =   7215
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnValor 
               Height          =   315
               Left            =   6525
               TabIndex        =   20
               Top             =   1440
               Width           =   1830
               _Version        =   65536
               _ExtentX        =   3228
               _ExtentY        =   556
               Calculator      =   "frmManIndicadores.frx":0F02
               Caption         =   "frmManIndicadores.frx":0F22
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManIndicadores.frx":0F8E
               Keys            =   "frmManIndicadores.frx":0FAC
               Spin            =   "frmManIndicadores.frx":0FF6
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
               MaxValue        =   999999999999999
               MinValue        =   -999999999999999
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
            Begin TrueOleDBList70.TDBCombo tdbcVariables 
               Height          =   300
               Left            =   1170
               TabIndex        =   13
               Tag             =   "enabled"
               Top             =   990
               Width           =   7185
               _ExtentX        =   12674
               _ExtentY        =   529
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               _DropdownWidth  =   15240
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
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
               _PropDict       =   $"frmManIndicadores.frx":101E
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
            Begin MSForms.CommandButton arbuBorrar 
               Height          =   390
               Left            =   8505
               TabIndex        =   12
               ToolTipText     =   "Limpiar formula y observacion"
               Top             =   270
               Width           =   1305
               Caption         =   " Limpiar"
               PicturePosition =   327683
               Size            =   "2302;688"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.CommandButton arbuAgregarOpe 
               Height          =   390
               Left            =   4140
               TabIndex        =   16
               ToolTipText     =   "Agregar operador a la formula"
               Top             =   1395
               Width           =   1305
               Caption         =   " Agregar"
               PicturePosition =   327683
               Size            =   "2302;688"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.CommandButton arbuAgregarVal 
               Height          =   390
               Left            =   8505
               TabIndex        =   18
               ToolTipText     =   "Agregar valor numerico a la formula"
               Top             =   1440
               Width           =   1305
               Caption         =   " Agregar"
               PicturePosition =   327683
               Size            =   "2302;688"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.CommandButton arbuAgregarVar 
               Height          =   390
               Left            =   8505
               TabIndex        =   14
               ToolTipText     =   "Agregar variable a la formula"
               Top             =   945
               Width           =   1305
               Caption         =   " Agregar"
               PicturePosition =   327683
               Size            =   "2302;688"
               Picture         =   "frmManIndicadores.frx":10A5
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
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
               Left            =   90
               TabIndex        =   35
               Top             =   2070
               Width           =   1020
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
               Left            =   5805
               TabIndex        =   17
               Top             =   1485
               Width           =   450
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Variables"
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
               TabIndex        =   34
               Top             =   990
               Width           =   810
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
               TabIndex        =   33
               Top             =   1485
               Width           =   810
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
               Left            =   90
               TabIndex        =   32
               Top             =   225
               Width           =   690
            End
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   8640
            TabIndex        =   5
            Tag             =   "_"
            Top             =   255
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "frmManIndicadores.frx":163F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManIndicadores.frx":16AB
            Key             =   "frmManIndicadores.frx":16C9
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
            AllowSpace      =   0
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   3
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
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   1395
            TabIndex        =   6
            Tag             =   "_"
            Top             =   615
            Width           =   8610
            _Version        =   65536
            _ExtentX        =   15187
            _ExtentY        =   556
            Caption         =   "frmManIndicadores.frx":171B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManIndicadores.frx":1787
            Key             =   "frmManIndicadores.frx":17A5
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
            MaxLength       =   250
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
         Begin TrueOleDBList70.TDBCombo tdbcTipo 
            Height          =   300
            Left            =   1395
            TabIndex        =   4
            Tag             =   "_"
            Top             =   255
            Width           =   3240
            _ExtentX        =   5715
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=847"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=767"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
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
            _PropDict       =   $"frmManIndicadores.frx":17F7
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
         Begin TDBNumber6Ctl.TDBNumber tdbnPorcMin 
            Height          =   315
            Left            =   5850
            TabIndex        =   9
            Tag             =   "_"
            Top             =   990
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   556
            Calculator      =   "frmManIndicadores.frx":187E
            Caption         =   "frmManIndicadores.frx":189E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManIndicadores.frx":190A
            Keys            =   "frmManIndicadores.frx":1928
            Spin            =   "frmManIndicadores.frx":1980
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
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
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
            MaxValueVT      =   1802698757
            MinValueVT      =   1769209861
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnPorcMax 
            Height          =   315
            Left            =   8460
            TabIndex        =   10
            Tag             =   "_"
            Top             =   990
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   556
            Calculator      =   "frmManIndicadores.frx":19A8
            Caption         =   "frmManIndicadores.frx":19C8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManIndicadores.frx":1A34
            Keys            =   "frmManIndicadores.frx":1A52
            Spin            =   "frmManIndicadores.frx":1AAA
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
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
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
            MaxValueVT      =   1802698757
            MinValueVT      =   1769209861
         End
         Begin TrueOleDBList70.TDBCombo tdbcUnidad 
            Height          =   300
            Left            =   1395
            TabIndex        =   7
            Tag             =   "_"
            Top             =   990
            Width           =   1845
            _ExtentX        =   3254
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=847"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=767"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
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
            _PropDict       =   $"frmManIndicadores.frx":1AD2
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
         Begin MSForms.CheckBox chkFlag 
            Height          =   285
            Left            =   3240
            TabIndex        =   8
            Top             =   1005
            Width           =   1590
            VariousPropertyBits=   746596379
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2805;503"
            Value           =   "0"
            Caption         =   "Razón Favorable"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidad"
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
            Left            =   270
            TabIndex        =   36
            Top             =   1035
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   225
            Index           =   3
            Left            =   285
            TabIndex        =   30
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Razón Min."
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
            Left            =   4905
            TabIndex        =   29
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            Left            =   285
            TabIndex        =   28
            Top             =   255
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
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
            Index           =   9
            Left            =   7785
            TabIndex        =   27
            Top             =   315
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Razón Max."
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
            Index           =   8
            Left            =   7515
            TabIndex        =   26
            Top             =   1035
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4140
         Left            =   -74820
         TabIndex        =   22
         Top             =   405
         Width           =   9900
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
            Height          =   255
            Left            =   360
            TabIndex        =   0
            Top             =   600
            Width           =   1095
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   2595
            Left            =   360
            TabIndex        =   3
            Top             =   1440
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   4577
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Tipo de Ratio"
            Columns(0).DataField=   "Ind_cTipoRatios"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre Tipo de Ratio"
            Columns(1).DataField=   "Tab_cDescripCampo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Código"
            Columns(2).DataField=   "Ind_cCodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Descripción"
            Columns(3).DataField=   "Ind_cDescripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   16
            Columns(4)._MaxComboItems=   5
            Columns(4).ValueItems(0)._DefaultItem=   0
            Columns(4).ValueItems(0).Value=   "1"
            Columns(4).ValueItems(0).Value.vt=   8
            Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(2)=   "7+vx7+vx7+vx7+vx7+tSpUoAlAhrtWPx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(3)=   "7+vx7+sYtSkAvSEAlACMvXvx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+trtWMAvSEA"
            Columns(4).ValueItems(0).DisplayValue(4)=   "xikApQAxnDHx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAnBAAzjEAxikArRAAlACl"
            Columns(4).ValueItems(0).DisplayValue(5)=   "xpTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tSpUoAzjEAxikA/2MAzjEAnAAAjADx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(6)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+sYtSkpzloA/2MA/2MAvSEAxikAlACMvXvx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(7)=   "7+vx7+vx7+vx7+sYxkIA/2MA/2NSpUpSpUoAxikApQAxnDHx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(8)=   "7+vx7+sArSEArSHx7+vx7+sArRgAxikAlAClxpTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(9)=   "7+vx7+vx7+sxtUIAxikAnAAAjADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(10)=   "7+sAtSEAxikAlACMvXvx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+tSpUoAxikp"
            Columns(4).ValueItems(0).DisplayValue(11)=   "rTkxtULx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sprUpa56UprTmMvXvx"
            Columns(4).ValueItems(0).DisplayValue(12)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sxtUIA1kKMvXvx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(13)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+ulxpTx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(0).DisplayValue(15)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+s="
            Columns(4).ValueItems(0).DisplayValue.vt=   9
            Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems(1)._DefaultItem=   0
            Columns(4).ValueItems(1).Value=   "0"
            Columns(4).ValueItems(1).Value.vt=   8
            Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAAMYDAABCTcYDAAAAAAAANgAAACgAAAAQAAAAEwAAAAEAGAAAAAAAkAMAAAAAAAAAAAAAAAAA"
            Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(2)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(3)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(4)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(5)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(6)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(7)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(8)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(9)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(10)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(11)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(12)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(13)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(15)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(16)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
            Columns(4).ValueItems(1).DisplayValue(17)=   "7+vx7+s="
            Columns(4).ValueItems(1).DisplayValue.vt=   9
            Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems.Count=   2
            Columns(4).Caption=   "Fav."
            Columns(4).DataField=   "Ind_cFlag"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Raz.Min"
            Columns(5).DataField=   "Ind_nPorceMin"
            Columns(5).NumberFormat=   "Standard"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Raz.Max"
            Columns(6).DataField=   "Ind_nPorceMax"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Unidad"
            Columns(7).DataField=   "Ind_cUnidad"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2884"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2805"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1085"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1005"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=529"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=5821"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5741"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=528"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=847"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=767"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1905"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1826"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=529"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1429"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1349"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=529"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=926"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=847"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=529"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   1
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(70)  =   "Named:id=33:Normal"
            _StyleDefs(71)  =   ":id=33,.parent=0"
            _StyleDefs(72)  =   "Named:id=34:Heading"
            _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   ":id=34,.wraptext=-1"
            _StyleDefs(75)  =   "Named:id=35:Footing"
            _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=36:Selected"
            _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=37:Caption"
            _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(81)  =   "Named:id=38:HighlightRow"
            _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=39:EvenRow"
            _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(85)  =   "Named:id=40:OddRow"
            _StyleDefs(86)  =   ":id=40,.parent=33"
            _StyleDefs(87)  =   "Named:id=41:RecordSelector"
            _StyleDefs(88)  =   ":id=41,.parent=34"
            _StyleDefs(89)  =   "Named:id=42:FilterBar"
            _StyleDefs(90)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   5925
            TabIndex        =   2
            Tag             =   "_"
            Top             =   825
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   556
            Caption         =   "frmManIndicadores.frx":1B59
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManIndicadores.frx":1BC5
            Key             =   "frmManIndicadores.frx":1BE3
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
            AllowSpace      =   0
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   250
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
         Begin TrueOleDBList70.TDBCombo tdbcTipoBus 
            Height          =   300
            Left            =   360
            TabIndex        =   1
            Tag             =   "_"
            Top             =   855
            Width           =   3615
            _ExtentX        =   6376
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=847"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=767"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
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
            _PropDict       =   $"frmManIndicadores.frx":1C35
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
            TabIndex        =   24
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Index           =   14
            Left            =   5895
            TabIndex        =   23
            Top             =   540
            Width           =   990
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":1CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":2096
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":2470
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":284A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":2C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":2FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":33D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":37B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   11025
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":47CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4926
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":4FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":5142
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":529C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   11025
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":53F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":5990
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":5F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":64C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":6A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":6FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":7592
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":7B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManIndicadores.frx":80C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
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
Attribute VB_Name = "frmManIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim Control As String
Dim cSepFormula As String
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

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

Public Function Replicar(cadena As String, veces As Integer) As String
    Dim i As Integer
    Replicar = ""
    For i = 1 To veces
        Replicar = Replicar & cadena
    Next i
End Function

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
            pSetFocus tdbtObserva
        Else
            tdbtObserva.SelText = cSepFormula + cOper + cSepFormula
            tdbtObserva = Trim(tdbtObserva)
            tdbtObserva.SelStart = Len(tdbtObserva)
            pSetFocus tdbtObserva
        End If
        
    End If
End If

End Sub

Private Sub arbuAgregarVal_Click()
If tdbnValor.Value <> 0 Then
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
End If
End Sub

Private Sub arbuAgregarVar_Click()
If tdbcVariables.Text <> "" Then
    If ValidaInsOper(4) = True Then
        tdbtFormula.SelText = cSepFormula + tdbcVariables.Columns(0) + cSepFormula
        tdbtFormula = Trim(tdbtFormula)
        tdbtFormula.SelStart = Len(tdbtFormula)
        pSetFocus tdbtFormula
        
        tdbtObserva.SelText = cSepFormula + tdbcVariables + cSepFormula
        tdbtObserva = Trim(tdbtObserva)
        tdbtObserva.SelStart = Len(tdbtObserva)
        pSetFocus tdbtObserva
        
    End If
End If
End Sub

Private Sub arbuBorrar_Click()
    tdbtFormula.Text = ""
    tdbtObserva.Text = ""
End Sub

Private Sub chkTipo_Click()
    Call FiltrarRecordSet
End Sub

Private Sub chkTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcTipoBus
End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 100
            .Height = Me.Height - .Top - 400
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 350
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 350
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        'Call Centrar_Objeto(Frame2, SSTCentroCosto, 200, 500)
        
        Frame2.Height = Frame1.Height + 100
        Frame2.Width = Frame1.Width + 100
        
        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
End Sub

Private Sub SSTCentroCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTCentroCosto.Tab = 0 Then pSetFocus chkTipo
    If SSTCentroCosto.Tab = 1 Then pSetFocus tdbcTipo
End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo
        Case 2: VerDatos
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
        Case 4: Borrar
        Case 5: Editar
        Case 6: Imprimir
        Case 7
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
            End If
    End Select
End Sub

Private Sub Borrar()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgCostos.Columns(0).Value) <> "" Then
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            Call CargaArregloMnt
            lArrMnt(0) = "ELIMINAR"                     ' Accion
            lArrMnt(2) = tdbgCostos.Columns(0).Value    ' Tipo
            lArrMnt(3) = tdbgCostos.Columns(2).Value    ' Codigo
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaIndicadores", lArrMnt(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Call CargaTabla
            Screen.MousePointer = vbDefault
            Mensajes "Registro ha sido eliminado", vbInformation
        End If
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

Private Sub VerDatos()

If Me.tdbgCostos.Columns(1).Value <> "" Then

    If lRegElim = False Then
        
        SSTCentroCosto.TabEnabled(1) = True
        SSTCentroCosto.TabEnabled(0) = False
        SSTCentroCosto.Tab = 1
        tbrOpciones.Buttons(1).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(2).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(5).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(7).Image = 8
        lTipoMnt = "EDITAR"
        Call AseguraControl(Me, False)
        Call ActivarBotones(False)
        tdbcOperador.Enabled = False
        tdbtObserva.Enabled = False
        tdbnValor.Enabled = False
        tdbcTipoBus.Enabled = True
            
        tdbnValor.Value = 0
        tdbtFormula.Text = ""
        tdbtObserva.Text = ""
        
        tdbcUnidad.Enabled = False
        tdbcUnidad.Locked = True
        
        chkFlag.Enabled = False
        
        tdbtFormula.Enabled = False
    Else
        lRegElim = False
    End If
    
    Call CargaDatosRegistro

Else
        Mensajes "No existe ningun registro seleccionado", vbInformation
End If
    
End Sub

Private Sub ActivarBotones(Valor As Boolean)
    arbuAgregarOpe.Enabled = Valor
    arbuAgregarVal.Enabled = Valor
    arbuAgregarVar.Enabled = Valor
    arbuBorrar.Enabled = Valor
End Sub

Private Sub Editar()
    Call CargaDatosRegistro
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        Call AseguraControl(Me, False)
        Call TabMantenimiento(True)
        Call ActivarBotones(True)
        tdbcTipo.Locked = True
        'tdbtCodigo.ReadOnly = True
        
        tdbcVariables.Locked = False
        tdbtDescripcion.ReadOnly = False
        tdbnPorcMin.ReadOnly = False
        tdbnPorcMax.ReadOnly = False

        tdbcVariables.Enabled = True
        tdbtDescripcion.Enabled = True
        tdbnPorcMin.Enabled = True
        tdbnPorcMax.Enabled = True
        tdbtCodigo.Enabled = True

        chkFlag.Enabled = True
        
        tdbcOperador.Enabled = True
        tdbtObserva.Enabled = True
        tdbnValor.Enabled = True
        
        tdbcTipoBus.Enabled = True
        tdbtFormula.Enabled = True
        
        tdbcUnidad.Enabled = True
        tdbcUnidad.Locked = False
    Else
        lRegElim = False
    End If
End Sub

Private Sub ManNuevo()
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call AseguraControl(Me, True)
    Call ActivarBotones(True)

    Call TabMantenimiento(True)
    
    tdbcOperador.Enabled = True
    tdbtObserva.Enabled = True
    tdbnValor.Enabled = True
    tdbcTipo.Enabled = True
    tdbcTipoBus.Enabled = True
    
    tdbnPorcMin.ReadOnly = False
    tdbnPorcMax.ReadOnly = False
    
    tdbcTipo.Locked = False
    tdbcTipoBus.Locked = False
    
    chkFlag.Enabled = True
    
    tdbtCodigo.ReadOnly = False
    tdbtDescripcion.ReadOnly = False
    tdbtObserva.Text = ""
    tdbnValor.Value = "0.00"
    tdbtFormula.Text = ""
    tdbcVariables.Locked = False
    
    tdbtFormula.Enabled = True
    
    tdbcUnidad.Enabled = True
    tdbcUnidad.Locked = False
    
    If chkTipo.Value = vbChecked Then
        tdbcTipo.BoundText = tdbcTipoBus.BoundText
    End If
    
    pSetFocus tdbcTipo
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    SSTCentroCosto.TabEnabled(1) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    If Valor = True Then SSTCentroCosto.Tab = 1
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
    If Valor = True Then
        tbrOpciones.Buttons(7).Image = 8
    Else
        tbrOpciones.Buttons(7).Image = 7
    End If
End Sub

Private Sub Cancelar()
    Call TabMantenimiento(False)
    tdbcTipoBus.Enabled = True
    pSetFocus tdbgCostos
End Sub

Private Sub Imprimir()
    Screen.MousePointer = vbHourglass
    Dim matriz_fecha(15) As Variant
    Dim Tipo As String
    
    matriz_fecha(0) = "@Accion;BUSCARTODOS;True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Ind_cTipoRatios;;True"
    matriz_fecha(3) = "@Ind_cCodigo;;True"
    matriz_fecha(4) = "@Ind_cDescripcion;;True"
    matriz_fecha(5) = "@Ind_nPorceMin;0;True"
    matriz_fecha(6) = "@Ind_nPorceMax;0;True"

    matriz_fecha(7) = "@Ind_cFormula;;True"
    matriz_fecha(8) = "@Ind_cObservacion;;True"

    matriz_fecha(9) = "@Ind_cEstado;;True"
    matriz_fecha(10) = "@Ind_cUserCrea;;True"
    matriz_fecha(11) = "@Ind_cUnidad;;True"
    matriz_fecha(12) = "@Ind_cFlag;;True"
    matriz_fecha(13) = "EmpresaNom;" & gsEmpresaNom & ";True"
    
    matriz_fecha(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(15) = "@RUC;" & "RUC : " & gsRUC & ";True"

    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptRatiosListado.rpt", crptToWindow, "Reporte de Analisis de Ratios", "", matriz_fecha(), formulas()
    
    
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub Grabar()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    If validarDatos = False Then Exit Sub
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaIndicadores", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    Call Cancelar
    CargaTabla
'    ' *** Buscar la Cuenta creado y posicionarse alli
'    Dim valor As Integer
'    valor = BuscarCadRs(tdbtCodigo, lrsTabla, 1)
'    If valor = 0 Then lrsTabla.MoveFirst
    ' ***
    FiltrarRecordSet
    Mensajes "Los datos se grabaron con exito...", vbInformation
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False

    If TextoLleno(tdbtCodigo, "Codigo") = False Then Exit Function
    If TextoLleno(tdbtDescripcion, "Descripcion") = False Then Exit Function
    If TextoLleno(tdbtFormula, "Formula") = False Then Exit Function
    If TextoLleno(tdbtObserva, "Observacion") = False Then Exit Function
    
    If Me.tdbnPorcMin.Value > Me.tdbnPorcMax Then
        MsgBox "El porcentaje minimo no debe pasar al porcentaje mayor", vbOKOnly + vbInformation, "Cuidado..."
        validarDatos = False
        Exit Function
    End If


    If ValidaFormula(tdbtFormula.Text, 6) = False Then
        validarDatos = False
        Exit Function
    End If

    validarDatos = True
End Function

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(12) As Variant
    lArrMnt(0) = lTipoMnt               ' Accion
    lArrMnt(1) = gsEmpresa              ' Empresa
    lArrMnt(2) = tdbcTipo.BoundText     ' Tipo de Ratios
    lArrMnt(3) = tdbtCodigo             ' Codigo
    lArrMnt(4) = tdbtDescripcion        ' Descripcion
    lArrMnt(5) = tdbnPorcMin            ' % Minimo
    lArrMnt(6) = tdbnPorcMax            ' % Maximo
    lArrMnt(7) = tdbtFormula            ' Formula
    lArrMnt(8) = tdbtObserva            ' Observacion
    lArrMnt(9) = "A"                    ' Estado
    lArrMnt(10) = gsUsuario             ' Usuario
    lArrMnt(11) = tdbcUnidad.BoundText  ' unidad
    
    If chkFlag.Value = True Then
        lArrMnt(12) = "1"
    Else
        lArrMnt(12) = "0"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
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

    Call Centrar_form(Me)
    
    ' *** Llenando las grillas y los combos
    Call LlenaCombos
    Call CargaTabla
    cSepFormula = " "
    ' *** Seteando algunas variables
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    
    tdbnValor.Value = "0.00"
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    SSTCentroCosto.Tab = 0
End Sub

Private Sub combo_operador()
    tdbcOperador.List(0) = " " + " ; " + "<NINGUNO>"
    '***** Operadores de Agrupación *****
    tdbcOperador.List(1) = "(" + " ; " + "ABRIR PARENTESIS"
    tdbcOperador.List(2) = ")" + " ; " + "CERRAR PARENTESIS"
    '***** Operadores Aritméticos *****
    tdbcOperador.List(3) = "+" + " ; " + "SUMAR"
    tdbcOperador.List(4) = "-" + " ; " + "RESTAR"
    tdbcOperador.List(5) = "*" + " ; " + "MULTIPLICAR"
    tdbcOperador.List(6) = "/" + " ; " + "DIVIDIR"
    
    On Error Resume Next
    tdbcOperador.ListIndex = 0
    
End Sub

Private Sub CargaVariables()
    Dim sqlSp As String
    
    sqlSp = "select Ppa_cTipoPlantilla + Ppa_cNumPlantilla AS Ppa_cNumPlantilla, Ppa_cNombre " & _
            "FROM CNA_TIPO_PLANTILLA WHERE Emp_cCodigo= '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' and Ppa_cNombre<>'' ORDER BY DBO.TRIMSQL(Ppa_cNombre)"
    
    LlenarComboAddItem tdbcVariables, sqlSp
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos  As String
    '------------------------------------------------------------------------------------------------
    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' and Tab_cTabla = '022' ORDER BY Tab_cCodigo"
    LlenarComboAddItem tdbcTipo, sqlcombos
    LlenarComboAddItem tdbcTipoBus, sqlcombos, True
    '------------------------------------------------------------------------------------------------
    sqlcombos = "select tab_ccodigo, tab_cdescripcampo FROM TABLA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' and Tab_cTabla = '052' ORDER BY Tab_cCodigo"
    LlenarComboAddItem tdbcUnidad, sqlcombos
    '------------------------------------------------------------------------------------------------
    Call combo_operador
    Call CargaVariables

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim registros As Integer
    
    registros = 0
    
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaIndicadores 'SEL_ALL', '" & gsEmpresa & "', '', '', '', 0, 0, '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then
       If lrsTabla.State = adStateOpen Then
        
            ' *** Llenar grilla con el RecordSet
            lrsTabla.Sort = "Ind_cTipoRatios,Ind_cCodigo, Ind_cDescripcion"
            tdbgCostos.DataSource = lrsTabla
            ' ***
            registros = lrsTabla.RecordCount
        End If
    End If
    
    
End Sub

Private Sub DesactivaBarraHerramientas(Valor As Boolean)
    tbrOpciones.Buttons(1).Enabled = Valor
    tbrOpciones.Buttons(2).Enabled = Valor
    tbrOpciones.Buttons(3).Enabled = Valor
    tbrOpciones.Buttons(4).Enabled = Valor
    tbrOpciones.Buttons(5).Enabled = Valor
    tbrOpciones.Buttons(6).Enabled = Valor
End Sub
Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    With tdbgCostos
        sqlSp = "spCn_GrabaIndicadores 'SEL_REG', '" & gsEmpresa & "', '" & .Columns(0).Value & "', '" & .Columns(2).Value & "', '', 0, 0, '', '','','' "
    End With
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro datos del registro seleccionado.", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    ' *** Asignando Datos a la Cuenta de Banco
    tdbcTipo.BoundText = CE(rsArreglo!Ind_cTipoRatios)
    tdbtCodigo = CE(rsArreglo!Ind_cCodigo)
    tdbtDescripcion = CE(rsArreglo!Ind_cDescripcion)
    tdbnPorcMin = NE(rsArreglo!Ind_nPorceMin)
    tdbnPorcMax = NE(rsArreglo!Ind_nPorceMax)
    tdbtFormula = CE(rsArreglo!Ind_cFormula)
    tdbtObserva = CE(rsArreglo!Ind_cObservacion)
    tdbcUnidad.BoundText = CE(rsArreglo!Ind_cUnidad)
    
    If NE(rsArreglo!Ind_cFlag) = 1 Then
        chkFlag.Value = vbChecked
    Else
        chkFlag.Value = vbUnchecked
    End If
    
    
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(1) As String
    Dim i As Integer
    If lrsTabla Is Nothing Then Exit Sub
    cadena = ""
    If chkTipo.Value = "1" Then filtros(0) = "Tab_cDescripCampo like '" & tdbcTipoBus.Text & "*'"
    If tdbtDescripcionBus <> "" Then filtros(1) = "Ind_cDescripcion like '*" & tdbtDescripcionBus & "*'"
    For i = 0 To 1
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    ' *** Filtrando segun campos
    If Trim(cadena) <> "" Then
        lrsTabla.Filter = cadena
    Else
        lrsTabla.Filter = 0
    End If
End Sub

Private Sub tdbcOperador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarOpe_Click
    End If
End Sub

Private Sub tdbcTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If tdbcTipo.Text <> "" Then
            tdbtCodigo.Text = BauscaCodigo(tdbcTipo.BoundText)
        End If
    
        pSendKeys "{tab}"
    End If
End Sub

Private Sub tdbcTipo_LostFocus()
    tdbtCodigo.Text = BauscaCodigo(tdbcTipo.BoundText)
End Sub

Private Sub tdbcTipo_SelChange(Cancel As Integer)
    If tdbcTipo.Text <> "" Then
        tdbtCodigo.Text = BauscaCodigo(tdbcTipo.BoundText)
    End If

End Sub

Private Sub tdbcTipoBus_ItemChange()
    Call FiltrarRecordSet
End Sub

Private Sub tdbcTipoBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtDescripcionBus
End If
End Sub

Private Sub tdbcVariables_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarVar_Click
    End If
End Sub

Private Sub tdbgCostos_DblClick()
    VerDatos
End Sub

Private Sub tdbgCostos_GotFocus()
tdbgCostos.HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgCostos_HeadClick(ByVal ColIndex As Integer)
If Not lrsTabla Is Nothing Then
    If lrsTabla.RecordCount > 0 Then
    
        lrsTabla.Sort = tdbgCostos.Columns(ColIndex).DataField
        tdbgCostos.DataSource = lrsTabla
        
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

Private Sub tdbnValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        arbuAgregarVal_Click
    End If
End Sub

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtCodigo = Replace(tdbtCodigo, "'", "")
       tdbtCodigo.SelStart = Len(tdbtCodigo)
    End If
End Sub

Private Sub tdbtCodigo_LostFocus()
    ' *** Verificar q codigo exista
    If SSTCentroCosto.TabEnabled(1) = True And lTipoMnt = "INSERTAR" Then
        If ExisteCodigo(tdbcTipo.BoundText, tdbtCodigo) = True Then
            Mensajes "Codigo ya existe. Verifique...", vbInformation
            pSetFocus tdbtCodigo
        End If
    End If
End Sub

Private Function BauscaCodigo(Tipo As String) As String
    ' *** Verificar q codigo exista
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaIndicadores 'BUSCACODIGO', '" & gsEmpresa & "', '" & Tipo & "', '', '', 0, 0, '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        BauscaCodigo = CE(rsArreglo!Codigo)
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function ExisteCodigo(Tipo As String, Valor As String) As Boolean
    ' *** Verificar q codigo exista
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    ExisteCodigo = False
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaIndicadores 'SEL_REG', '" & gsEmpresa & "', '" & Tipo & "', '" & Valor & "', '', 0, 0, '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteCodigo = True
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function existeCtaBalance(cadena As String) As String
    Dim sqlver As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    existeCtaBalance = ""
    sqlver = "SELECT Ind_cDescripcion From dbo.CNT_CUENTA_INDI WHERE Emp_cCodigo = '" & gsEmpresa & "' "
    sqlver = sqlver + "AND Ind_cCodCuenta = '" & cadena & "'"
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "Codigo no existe. Verificar", vbInformation
        Exit Function
    End If
    existeCtaBalance = rsArreglo(0).Value
    Call CerrarRecordSet(rsArreglo)
End Function

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

Private Sub tdbtFormula_KeyPress(KeyAscii As Integer)
Dim nPos As Integer

If KeyAscii = 8 And CE(tdbtFormula.Text) <> "" Then
' Presiona BACKSPACE
    nPos = InStrRev(tdbtFormula, cSepFormula, Len(tdbtFormula))
    If nPos >= 0 Then tdbtFormula = Trim(Left(tdbtFormula, nPos))
    tdbtFormula.SelStart = Len(tdbtFormula)
End If
End Sub

Private Sub tdbtObserva_Change()
    If CE(tdbtFormula.Text) = "" Then
       arbuBorrar_Click
    End If

End Sub
