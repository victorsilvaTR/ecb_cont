VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmManCentroCostoNiv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Centro de Costo"
   ClientHeight    =   4860
   ClientLeft      =   4410
   ClientTop       =   2430
   ClientWidth     =   7890
   Icon            =   "frmManCentroCostoNiv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   7890
   Begin TabDlg.SSTab sstDatos 
      Height          =   4410
      Left            =   -15
      TabIndex        =   10
      Top             =   390
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   7779
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Listado de Centro de Costos"
      TabPicture(0)   =   "frmManCentroCostoNiv.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos de Centro de Costos"
      TabPicture(1)   =   "frmManCentroCostoNiv.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraNivel"
      Tab(1).Control(1)=   "FrameDatos"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraNivel 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   420
         Left            =   -74190
         TabIndex        =   0
         Top             =   495
         Width           =   6315
         Begin MSDataListLib.DataCombo tdbcNivel 
            Height          =   300
            Left            =   1530
            TabIndex        =   3
            Top             =   0
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblNivel 
            AutoSize        =   -1  'True
            Caption         =   "N° de Niveles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   21
            Top             =   45
            Width           =   1185
         End
      End
      Begin VB.Frame FrameDatos 
         Height          =   3345
         Left            =   -74235
         TabIndex        =   13
         Top             =   855
         Width           =   6360
         Begin VB.Frame fraDetalle 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1500
            Left            =   225
            TabIndex        =   14
            Top             =   1710
            Width           =   5910
            Begin TDBText6Ctl.TDBText tdbtNombre 
               Height          =   330
               Left            =   1350
               TabIndex        =   8
               Top             =   495
               Width           =   4425
               _Version        =   65536
               _ExtentX        =   7805
               _ExtentY        =   582
               Caption         =   "frmManCentroCostoNiv.frx":0F02
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManCentroCostoNiv.frx":0F6E
               Key             =   "frmManCentroCostoNiv.frx":0F8C
               BackColor       =   16777215
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
               MaxLength       =   0
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
            Begin TDBText6Ctl.TDBText tdbtNombreCorto 
               Height          =   315
               Left            =   1350
               TabIndex        =   9
               Top             =   990
               Visible         =   0   'False
               Width           =   1800
               _Version        =   65536
               _ExtentX        =   3175
               _ExtentY        =   556
               Caption         =   "frmManCentroCostoNiv.frx":0FE2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManCentroCostoNiv.frx":104E
               Key             =   "frmManCentroCostoNiv.frx":106C
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
            Begin TDBText6Ctl.TDBText tdbtCodigo 
               Height          =   330
               Left            =   1350
               TabIndex        =   7
               Top             =   45
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Caption         =   "frmManCentroCostoNiv.frx":10C2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManCentroCostoNiv.frx":112E
               Key             =   "frmManCentroCostoNiv.frx":114C
               BackColor       =   16777215
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
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
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
            Begin VB.Label lblCodigo 
               Caption         =   "Código"
               Height          =   195
               Left            =   0
               TabIndex        =   17
               Top             =   90
               Width           =   1305
            End
            Begin VB.Label Label2 
               Caption         =   "Descripción"
               Height          =   315
               Left            =   0
               TabIndex        =   16
               Top             =   585
               Width           =   1170
            End
            Begin VB.Label Label3 
               Caption         =   "Descripcion Corta"
               Height          =   225
               Left            =   0
               TabIndex        =   15
               Top             =   1080
               Visible         =   0   'False
               Width           =   1350
            End
         End
         Begin VB.Frame fraCombos 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1410
            Left            =   225
            TabIndex        =   18
            Top             =   270
            Width           =   5820
            Begin TrueOleDBList70.TDBCombo tdbdSegNivel 
               Height          =   315
               Left            =   1350
               TabIndex        =   5
               Top             =   495
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   556
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=556"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
               HeadLines       =   0
               FootLines       =   1
               RowDividerStyle =   0
               Caption         =   ""
               EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               LayoutName      =   ""
               LayoutFileName  =   ""
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTips        =   0
               EditHeight      =   315.213
               AutoSize        =   -1  'True
               GapHeight       =   30.047
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
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               AddItemSeparator=   ";"
               _PropDict       =   $"frmManCentroCostoNiv.frx":11A2
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
            Begin TrueOleDBList70.TDBCombo tdbdPrimNivel 
               Height          =   315
               Left            =   1350
               TabIndex        =   4
               Top             =   0
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   556
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
               HeadLines       =   0
               FootLines       =   1
               RowDividerStyle =   0
               Caption         =   ""
               EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               LayoutName      =   ""
               LayoutFileName  =   ""
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTips        =   0
               EditHeight      =   315.213
               AutoSize        =   -1  'True
               GapHeight       =   30.047
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
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               AddItemSeparator=   ";"
               _PropDict       =   $"frmManCentroCostoNiv.frx":1229
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
            Begin TrueOleDBList70.TDBCombo tdbdTerNivel 
               Height          =   315
               Left            =   1350
               TabIndex        =   6
               Top             =   945
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   556
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
               HeadLines       =   0
               FootLines       =   1
               RowDividerStyle =   0
               Caption         =   ""
               EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               LayoutName      =   ""
               LayoutFileName  =   ""
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTips        =   0
               EditHeight      =   315.213
               AutoSize        =   -1  'True
               GapHeight       =   30.047
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
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
               AddItemSeparator=   ";"
               _PropDict       =   $"frmManCentroCostoNiv.frx":12B0
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
            Begin VB.Label lblTerNivel 
               Caption         =   "3er Nivel"
               Height          =   315
               Left            =   0
               TabIndex        =   22
               Top             =   990
               Width           =   975
            End
            Begin VB.Label lblPriNivel 
               Caption         =   "1er Nivel"
               Height          =   315
               Left            =   0
               TabIndex        =   20
               Top             =   45
               Width           =   975
            End
            Begin VB.Label lblSegNivel 
               Caption         =   "2do Nivel"
               Height          =   315
               Left            =   0
               TabIndex        =   19
               Top             =   540
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3855
         Left            =   488
         TabIndex        =   11
         Top             =   465
         Width           =   6915
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   1545
            TabIndex        =   1
            Tag             =   "ReadOnly"
            Top             =   285
            Width           =   4890
            _Version        =   65536
            _ExtentX        =   8625
            _ExtentY        =   556
            Caption         =   "frmManCentroCostoNiv.frx":1337
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManCentroCostoNiv.frx":13A3
            Key             =   "frmManCentroCostoNiv.frx":13C1
            BackColor       =   16777215
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
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
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
         Begin TrueOleDBGrid70.TDBGrid tdbgListado 
            Height          =   3060
            Left            =   360
            TabIndex        =   2
            Top             =   675
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5398
            _LayoutType     =   4
            _RowHeight      =   15
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "Cos_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "Cos_cDescripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1482"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1402"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8334"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8255"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H80000002&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.valignment=2"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.alignment=3"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HCA570B&,.bold=-1,.fontsize=825"
            _StyleDefs(28)  =   ":id=14,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.valignment=0"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.alignment=3"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(47)  =   "Named:id=33:Normal"
            _StyleDefs(48)  =   ":id=33,.parent=0"
            _StyleDefs(49)  =   "Named:id=34:Heading"
            _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   ":id=34,.wraptext=-1"
            _StyleDefs(52)  =   "Named:id=35:Footing"
            _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(54)  =   "Named:id=36:Selected"
            _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(56)  =   "Named:id=37:Caption"
            _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(58)  =   "Named:id=38:HighlightRow"
            _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(60)  =   "Named:id=39:EvenRow"
            _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(62)  =   "Named:id=40:OddRow"
            _StyleDefs(63)  =   ":id=40,.parent=33"
            _StyleDefs(64)  =   "Named:id=41:RecordSelector"
            _StyleDefs(65)  =   ":id=41,.parent=34"
            _StyleDefs(66)  =   "Named:id=42:FilterBar"
            _StyleDefs(67)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Caption         =   "Descripcion"
            Height          =   270
            Left            =   375
            TabIndex        =   12
            Top             =   345
            Width           =   1260
         End
      End
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   5085
      Top             =   0
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
            Picture         =   "frmManCentroCostoNiv.frx":1417
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1571
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":16CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1825
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":197F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1AD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1D8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":1EE7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   4455
      Top             =   0
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
            Picture         =   "frmManCentroCostoNiv.frx":2041
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":25DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":2B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":310F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":36A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":3C43
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":41DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":4777
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManCentroCostoNiv.frx":4D11
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   23
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
Attribute VB_Name = "frmManCentroCostoNiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmManCentroCostoNiv
'    Project    : Contabilidad
'
'    Description: Formulario de mantenimiento de centros de costos
'--------------------------------------------------------------------------------
Option Explicit
Dim flgModificado As Boolean    ' Utilizado para grabar o cancelar sin advertencia
Dim flgModoEdit As Boolean      ' Utilizado para evitar cambiar el flag Modificado mientras se cargan los datos
Dim rs_areas As New ADODB.Recordset
Dim flgNuevo As Boolean
Public sFormulario As String

Dim gsGrupo As String
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grupo
' Description:       Propiedad de seteo de grupo
'
' Parameters :       Grupo (String)
'--------------------------------------------------------------------------------
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el formulario
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim respuesta As String
Select Case KeyCode
    Case vbKeyReturn:
        If TypeOf Me.ActiveControl Is ComboBox Or _
            TypeOf Me.ActiveControl Is DataCombo Or _
            TypeOf Me.ActiveControl Is TextBox Or _
            TypeOf Me.ActiveControl Is CheckBox Or _
            TypeOf Me.ActiveControl Is OptionButton Then
                Call EnterTab(KeyCode): Exit Sub
        End If
    Case vbKeyBack, vbKeyClear
        If TypeOf Me.ActiveControl Is DataCombo Or _
            TypeOf Me.ActiveControl Is TDBCombo Then
            Me.ActiveControl.BoundText = "": Exit Sub
        End If
        If TypeOf Me.ActiveControl Is ComboBox Then
            Me.ActiveControl.ListIndex = 0: Exit Sub
        End If
    Case vbKeyEscape:
            If sstDatos.TabEnabled(0) = True Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                
                If respuesta = vbYes Then Call Cancelar
            End If
        
    Case vbKeyF2: If tbrOpciones.Buttons(1).Enabled Then nuevo
    Case vbKeyF3: If tbrOpciones.Buttons(2).Enabled Then Call VerDatos(False)
    Case vbKeyF4: If tbrOpciones.Buttons(3).Enabled Then Grabar
    Case vbKeyF5: If tbrOpciones.Buttons(4).Enabled Then Eliminar
    Case vbKeyF6: If tbrOpciones.Buttons(5).Enabled Then Call VerDatos(True)
    Case vbKeyF7: If tbrOpciones.Buttons(5).Enabled Then pPrint
End Select


End Sub

Private Sub BotonesVisibles(bNuevo As Boolean, bConsulta As Boolean, bGrabar As Boolean, bEliminar As Boolean, bEditar As Boolean, bImprimir As Boolean, bSalir As Boolean)
    tbrOpciones.Buttons(1).Enabled = bNuevo
    tbrOpciones.Buttons(2).Enabled = bConsulta
    tbrOpciones.Buttons(3).Enabled = bGrabar
    tbrOpciones.Buttons(4).Enabled = bEliminar
    tbrOpciones.Buttons(5).Enabled = bEditar
    tbrOpciones.Buttons(6).Enabled = bImprimir
    tbrOpciones.Buttons(7).Enabled = bSalir
    
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    Me.Tag = frmMDIConta.mnuCentroCostos.Tag
    
    Call CargarListado
    '-----------------------------------------
    Call LlenaComboNiveles
    '-------------------------------------------
    Call CargarPrimerNivel
    
    'SeteaBarraHerramientas Me.ToolBar_ECB1, gsGrupo
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaComboNiveles
' Description:       Procedimiento que llena los nivlees de los combos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaComboNiveles()
    Dim sNivel As String
    sNivel = fRetornaValor("spCNT_CONFIG_LIBROS 'BUSCARNIVEL','" & gsEmpresa & "','','','','','','','','',0,'','','','','','','','','','','','','','','','" & gsAnio & "'")
    
    Dim lrsNivel As New ADODB.Recordset
    With lrsNivel
        .CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
        .CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
        .LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
        .Fields.Append "CODIGO", adChar, 1
        .Fields.Append "DESCRIPCION", adVarChar, 50
        .Open
        .AddNew: .Fields("CODIGO") = " ": .Fields("DESCRIPCION") = "<Seleccione un nivel>"
        
        If sNivel = "P" Or sNivel = "S" Or sNivel = "T" Or sNivel = "C" Then
            .AddNew: .Fields("CODIGO") = "P": .Fields("DESCRIPCION") = "PRIMER NIVEL"
        End If
        
        If sNivel = "S" Or sNivel = "T" Or sNivel = "C" Then
            .AddNew: .Fields("CODIGO") = "S": .Fields("DESCRIPCION") = "SEGUNDO NIVEL"
        End If
        
        If sNivel = "T" Or sNivel = "C" Then
        .AddNew: .Fields("CODIGO") = "T": .Fields("DESCRIPCION") = "TERCER NIVEL"
        End If
        
        If sNivel = "C" Then
        .AddNew: .Fields("CODIGO") = "C": .Fields("DESCRIPCION") = "CUARTO NIVEL"
        End If
        
        .Update
    End With
    
    Set tdbcNivel.RowSource = lrsNivel
    tdbcNivel.ListField = "DESCRIPCION"
    tdbcNivel.BoundColumn = "CODIGO"
    


End Sub

Private Sub Form_Resize()
On Error GoTo errHand
'    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
'    Call Form_Maximizar(Me)

' Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

    If Me.WindowState <> vbMinimized Then
        '*** CENTRAR TITULO
        'ToolBar_ECB1.Width = Me.Width
'        lblTitulo.Width = Me.Width - ToolBar_ECB1.Width
'        lblTitulo.Alignment = vbCenter
        '*** REDIMENSIONAR SST
        sstDatos.Width = Me.Width - sstDatos.Left + 15 - 100
        sstDatos.Height = Me.Height - sstDatos.Top + 15
        '*** REDIMENSIONAR FRAME PRINCIPAL
        Frame1.Width = sstDatos.Width - IIf(sstDatos.TabOrientation = ssTabOrientationLeft, sstDatos.TabHeight, 0) - 700
        Frame1.Height = sstDatos.Height - IIf(sstDatos.TabOrientation = ssTabOrientationTop, sstDatos.TabHeight, 0) - 700
       
        '*** REDIMENSIONAR CUADRICULA DE LISTADO
        tdbgListado.Width = Frame1.Width - tdbgListado.Left - 300
        tdbgListado.Height = Frame1.Height - tdbgListado.Top - 200
        '*** REDIMENSIONAR DETALLE
        FrameDatos.Height = sstDatos.Height - IIf(sstDatos.TabOrientation = ssTabOrientationTop, sstDatos.TabHeight, 0) - 1000
        FrameDatos.Width = sstDatos.Width - IIf(sstDatos.TabOrientation = ssTabOrientationLeft, sstDatos.TabHeight, 0) - 1000
    End If
Exit Sub
errHand:
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al salir del formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Dim nRpta As VbMsgBoxResult
    nRpta = vbYes
    If flgModificado Then nRpta = Mensajes("¿Desea cancelar los cambios realizados?", vbQuestion + vbYesNo + vbDefaultButton2)
    If nRpta = vbNo Then Cancel = 1: Exit Sub
    
    Set tdbgListado.DataSource = Nothing
    Call CerrarRecordSet(rs_areas)
    
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargarListado
' Description:       Procedimiento que llena los niveles de centro de costos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CargarListado()
    Dim gtxtSQL As String
    gtxtSQL = "spCNT_CENTRO_COSTO 'BUSCAR_NIVEL_CONTA', '" & gsEmpresa & "', '" & gsAnio & "'"
    Set tdbgListado.DataSource = Nothing
    Call CerrarRecordSet(rs_areas)
    Set rs_areas = fRetornaRS(gtxtSQL)
    Set tdbgListado.DataSource = rs_areas
    If GetRsRecordCount(rs_areas) = 0 Then
        'ToolBar_ECB1.usrButtonsEnabled = "1" & String(7, "0")
    Else
        'ToolBar_ECB1.usrButtonsEnabled = String(8, "1")
        Call SetRsBookMarkOfCol(tdbtCodigo, rs_areas, 0)
    End If
    sstDatos.TabEnabled(0) = True
    sstDatos.TabEnabled(1) = False
    sstDatos.Tab = 0
    Call pSetFocus(tdbgListado)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargarPrimerNivel
' Description:       Procedimiento que llena el primer nivel
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CargarPrimerNivel()
    Dim sql As String
    sql = "spCNT_CENTRO_COSTO @Accion='BUSCA_PRI_NIVEL', @Emp_cCodigo='" & gsEmpresa & "',@Pan_cAnio='" & gsAnio & "'"
    LlenarComboAddItem tdbdPrimNivel, sql, True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargarSegundoNivel
' Description:       Procedimiento que llena el segundo nivel
'
' Parameters :       sPrimerNivel (String)
'--------------------------------------------------------------------------------
Private Sub CargarSegundoNivel(sPrimerNivel As String)
    Dim sql As String
    sql = "spCNT_CENTRO_COSTO @Accion='BUSCA_SEG_NIVEL', @Emp_cCodigo='" & gsEmpresa & "',@Pan_cAnio='" & gsAnio & "',@Cos_nPriNivel='" & sPrimerNivel & "'"
    LlenarComboAddItem tdbdSegNivel, sql, True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargarTercerNivel
' Description:       Procedimiento que llena el tercer nivel
'
' Parameters :       sPrimerNivel (String)
'                    sSegundoNivel (String)
'--------------------------------------------------------------------------------
Private Sub CargarTercerNivel(sPrimerNivel As String, sSegundoNivel As String)
    Dim sql As String
    sql = "spCNT_CENTRO_COSTO @Accion='BUSCA_TER_NIVEL', @Emp_cCodigo='" & gsEmpresa & "',@Pan_cAnio='" & gsAnio & "',@Cos_nPriNivel='" & sPrimerNivel & "',@Cos_nSegNivel='" & sSegundoNivel & "'"
    LlenarComboAddItem tdbdTerNivel, sql, True
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcNivel_ItemChange
' Description:       Evento que se ejecuta al cambiar elitem del nivel
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbcNivel_ItemChange()

End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: nuevo
                Call CancelarToolbar
        Case 2: VerDatos False
                Call BotonesVisibles(False, False, False, False, False, True, True)
        Case 3: Grabar
                
        Case 4: Eliminar
                Call BotonesVisibles(True, True, False, True, True, True, True)
        Case 5: VerDatos True
                If rs_areas.State = 0 Then Exit Sub
                Call CancelarToolbar
        Case 6: pPrint
        Case 7:
                If sstDatos.Tab = 0 Then
                    Unload Me
                Else
                    Cancelar
                    Call BotonesVisibles(True, True, False, True, True, True, True)
                End If
                
        Case 8: Unload Me
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcNivel_Validate
' Description:       Evento que se ejecuta al validar el primer nivel seleccionado
'
' Parameters :       Cancel (Boolean)
'--------------------------------------------------------------------------------
Private Sub tdbcNivel_Validate(Cancel As Boolean)
        tdbdPrimNivel.BoundText = ""
        tdbdSegNivel.BoundText = ""
        tdbdTerNivel.BoundText = ""
        
        Select Case tdbcNivel.BoundText
            Case "P"
                    ActivarControl tdbdPrimNivel, False
                    ActivarControl tdbdSegNivel, False
                    ActivarControl tdbdTerNivel, False
                    
                    ActivarControl tdbtNombre, True

            Case "S"
                    ActivarControl tdbdPrimNivel, True
                    ActivarControl tdbdSegNivel, False
                    ActivarControl tdbdTerNivel, False
                    
                    ActivarControl tdbtNombre, True

            Case "T"
                    ActivarControl tdbdPrimNivel, True
                    ActivarControl tdbdSegNivel, True
                    ActivarControl tdbdTerNivel, False
                    
                    ActivarControl tdbtNombre, True
                    
            Case "C"
                    ActivarControl tdbdPrimNivel, True
                    ActivarControl tdbdSegNivel, True
                    ActivarControl tdbdTerNivel, True
                    
                    ActivarControl tdbtNombre, True
                    
        End Select
        
        DoEvents
        
        Select Case tdbcNivel.BoundText
            Case "P"
                    tdbtCodigo.Text = CE(BuscaCodigo(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText, tdbdTerNivel.BoundText))
            Case "S"
                    Call CargarPrimerNivel

            Case "T"
                    Call CargarPrimerNivel
                    Call CargarSegundoNivel(tdbcNivel.BoundText)
                    
            Case "C"
                    Call CargarPrimerNivel
                    Call CargarSegundoNivel(tdbcNivel.BoundText)
                    Call CargarTercerNivel(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText)
                    
        End Select
        
        
        
        

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbdPrimNivel_ItemChange
' Description:       Evento que se ejecuta al cambiar el item del primer nivel
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbdPrimNivel_ItemChange()
    If tdbcNivel.BoundText = "T" Or tdbcNivel.BoundText = "C" Then
        Call CargarSegundoNivel(tdbdPrimNivel.BoundText)
    End If
    
    If tdbdPrimNivel.BoundText <> "" And tdbcNivel.BoundText = "S" Then
        tdbtCodigo.Text = CE(BuscaCodigo(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText, tdbdTerNivel.BoundText))
    Else
        tdbtCodigo.Text = ""
    End If
    
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbdSegNivel_ItemChange
' Description:       Evento que se ejecuta al cambiar el item del segundo nivel
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbdSegNivel_ItemChange()

    If tdbcNivel.BoundText = "C" Then
        Call CargarTercerNivel(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText)
    End If


    If tdbdPrimNivel.BoundText <> "" And tdbdSegNivel.BoundText <> "" And tdbcNivel.BoundText = "T" Then
        tdbtCodigo.Text = CE(BuscaCodigo(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText, tdbdTerNivel.BoundText))
    Else
        tdbtCodigo.Text = ""
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbdTerNivel_ItemChange
' Description:       Evento que se ejecuta al cambiar el item del tercer nivel
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbdTerNivel_ItemChange()

    If tdbdPrimNivel.BoundText <> "" And tdbdSegNivel.BoundText <> "" And tdbdTerNivel.BoundText <> "" And tdbcNivel.BoundText = "C" Then
        tdbtCodigo.Text = CE(BuscaCodigo(tdbdPrimNivel.BoundText, tdbdSegNivel.BoundText, tdbdTerNivel.BoundText))
    Else
        tdbtCodigo.Text = ""
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodigo_Change
' Description:       Evento que se ejecuta al canbiar el codigo de centro de costo
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCodigo_Change()
    If flgModoEdit Then flgModificado = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       BuscaCodigos
' Description:       Funcion que retorna si se encontro el codigo ingresado
'
' Parameters :       sCodigo (String)
'                    sCodigoInterno (Variant)
'--------------------------------------------------------------------------------
Private Function BuscaCodigos(sCodigo As String, sCodigoInterno) As Boolean
    On Error GoTo serror
    Dim i As Integer
    Dim Pos As Integer
    BuscaCodigos = False
    Pos = rs_areas.AbsolutePosition
    rs_areas.MoveFirst
    Do While Not rs_areas.EOF
        If CE(rs_areas.Fields("Cos_cCodigo")) = sCodigo And CE(rs_areas.Fields("Cot_cCodArea")) <> sCodigoInterno Then
            Mensajes "El codigo ingresado fue encontrado en el concepto " & Salto(1) & "Codigo " & CE(rs_areas.Fields("Cot_cCodArea")) & " : " & CE(rs_areas.Fields("Cot_cDescripLarga"))
            rs_areas.AbsolutePosition = Pos
            BuscaCodigos = True
            Exit Function
        End If
        rs_areas.MoveNext
    Loop
    rs_areas.AbsolutePosition = Pos
    Exit Function
serror:
    BuscaCodigos = False
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtDescripcion_Change
' Description:       Evento que se ejecuta al cambiar la descripcion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtDescripcion_Change()
    Dim cad As String
    If Trim(tdbtDescripcion) = "" Then
        rs_areas.Filter = 0
    Else
        cad = "CoS_cDescripCION LIKE '%" & Trim(tdbtDescripcion.Text) & "%'"
        rs_areas.Filter = cad
    End If
    Set tdbgListado.DataSource = rs_areas
'    If rs_areas.RecordCount = 0 Then
'        ToolBar_ECB1.usrButtonsEnabled = "1" & String(7, "0")
'    Else
'        ToolBar_ECB1.usrButtonsEnabled = String(8, "1")
'    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombre_Change
' Description:       Evento que se ejecuta al cambiar el nombre
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNombre_Change()
    If flgModoEdit Then flgModificado = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombre_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el nombre
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNombre_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "ñ" Then KeyAscii = Asc("Ñ")
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombreCorto_Change
' Description:       Evento que se ejecuta al cambiar el nombre corto
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNombreCorto_Change()
    If flgModoEdit Then flgModificado = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombreCorto_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el nombre corto
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNombreCorto_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "ñ" Then KeyAscii = Asc("Ñ")
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pPrint
' Description:       Procedimiento que imprime el reporte
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub pPrint()

    Dim matriz_fecha(5) As Variant
            matriz_fecha(0) = "@Accion;BUSCAR_NIVEL_CONTA;True"
            matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
            matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
            matriz_fecha(3) = "NombreEmp;" & gsEmpresaNom & ";True"
            matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
            matriz_fecha(5) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptCentroCosto.rpt", crptToWindow, "Reporte por Centro de Costo", "", matriz_fecha(), formulas()

End Sub




'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Cancelar
' Description:       Procedimiento que cancela la transaccion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Cancelar()
Dim nRpta As VbMsgBoxResult
    nRpta = vbYes
    'ToolBar_ECB1.usrNotChange = True     ' No cambia los botones
    If flgModificado Then nRpta = Mensajes("¿Desea cancelar los cambios realizados?", vbQuestion + vbYesNo + vbDefaultButton2)
    If nRpta = vbYes Then
        sstDatos.TabEnabled(0) = True
        sstDatos.TabEnabled(1) = False
        sstDatos.Tab = 0
        Call pSetFocus(tdbgListado)
        ' LIMPIO EL FLAG
        flgModificado = False
        'ToolBar_ECB1.usrNotChange = False
    End If
    
    Call BotonesVisibles(True, True, False, True, True, True, True)

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       BuscaCodigo
' Description:       Funcion que retorna el siguiente codigo de centro de costo
'
' Parameters :       sCodPriNivel (String)
'                    sCodSegNivel (String)
'                    sCodTerNivel (String)
'--------------------------------------------------------------------------------
Private Function BuscaCodigo(sCodPriNivel As String, sCodSegNivel As String, sCodTerNivel As String) As String
    Dim sql As String
    sql = "spCNT_CENTRO_COSTO @Accion='SIGUIENTECODIGO', @Emp_cCodigo='" & gsEmpresa & "',@Pan_cAnio='" & gsAnio & "',@Cos_nPriNivel='" & sCodPriNivel & "' ,@Cos_nSegNivel='" & sCodSegNivel & "', @Cos_nTerNivel='" & sCodTerNivel & "'"
    BuscaCodigo = fRetornaValor(sql)
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       nuevo
' Description:       Procedimiento que permite crear nuevos centros de costos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub nuevo()
    Call BotonesVisibles(False, False, True, False, False, True, True)
    Call OcultarCampos(False)
    
    flgNuevo = True
    ' DESACTIVA MODO DE EDICION
    flgModoEdit = False
    Call LimpiaTexto(Me)
    tdbtNombre.Text = ""
    tdbtCodigo.Text = ""
    sstDatos.TabEnabled(0) = False
    sstDatos.TabEnabled(1) = True
    sstDatos.Tab = 1
    FrameDatos.Enabled = True
    Call pSetFocus(tdbtNombre)
    ' ACTIVA MODO DE EDICION
    flgModoEdit = True
    tdbcNivel.BoundText = ""
    tdbdPrimNivel.BoundText = ""
    tdbdSegNivel.BoundText = ""
    tdbdTerNivel.BoundText = ""
    
    pSetFocus tdbcNivel
    'Call CargarPrimerNivel
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       BuscaDescrip
' Description:       Procedimiento que busca la descripcion de codigo
'
' Parameters :       sCodigo (String)
'                    scadena (String)
'--------------------------------------------------------------------------------
Private Function BuscaDescrip(sCodigo As String, scadena As String) As String

    'If Right(sCodigo, 4) = "0000" Then BuscaDescrip = scadena
    'If Right(sCodigo, 4) <> "0000" And Right(sCodigo, 2) = "00" Then BuscaDescrip = Right(scadena, Len(scadena) - 10)
    'If Right(sCodigo, 4) <> "0000" And Right(sCodigo, 2) <> "00" Then BuscaDescrip = Right(scadena, Len(scadena) - 12)
    
    BuscaDescrip = scadena
    BuscaDescrip = CE(BuscaDescrip)
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       OcultarCampos
' Description:       Procedimiento que oculta los frame segun el valor del parametro
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub OcultarCampos(bValor As Boolean)
    If bValor = True Then
        fraNivel.Visible = False
        fraCombos.Visible = False
        fraDetalle.Visible = True
        fraDetalle.Top = 270
    Else
        fraNivel.Visible = True
        fraCombos.Visible = True
        fraDetalle.Visible = True
        fraDetalle.Top = 1710
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       VerDatos
' Description:       Prcedimiento que permite ver los datos del centro de costo pero enconsulta
'
' Parameters :       bEdit (Boolean)
'--------------------------------------------------------------------------------
Private Sub VerDatos(ByVal bEdit As Boolean)
If rs_areas.State = 0 Then Exit Sub
    flgNuevo = False
    ' DESACTIVA MODO DE EDICION
    flgModoEdit = False
    
    Call OcultarCampos(True)
    
    
    If bEdit = True Then
        Call BotonesVisibles(False, False, True, False, False, True, True)
    Else
        Call BotonesVisibles(False, False, False, False, False, True, True)
    End If
    
    tdbtCodigo.Text = CE(rs_areas.Fields("Cos_cCodigo"))
    tdbtNombre.Text = BuscaDescrip(tdbgListado.Columns(0), tdbgListado.Columns(1))
    tdbtNombreCorto.Text = CE(rs_areas.Fields("Cot_cDescripCorta"))
    
    If CE(rs_areas.Fields("Nivel0")) = "P" Then
        'lblCodCC.Visible = True
        ActivarControl tdbtNombre, True
    Else
        'lblCodCC.Visible = False
        ActivarControl tdbtNombre, True
    End If
    '-------------------
    sstDatos.TabEnabled(0) = False
    sstDatos.TabEnabled(1) = True
    sstDatos.Tab = 1
    FrameDatos.Enabled = bEdit
    Call pSetFocus(tdbtNombre)
    ' ACTIVA MODO DE EDICION
    flgModoEdit = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ValidaDatos
' Description:       Procedimiento que valida los datosdel centro de costo antes de ser grabado
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If flgNuevo Then
        If tdbdPrimNivel.Enabled = True And tdbdPrimNivel.BoundText = "" Then Mensajes "Seleccione el 1er nivel de la lista": Call pSetFocus(tdbdPrimNivel): CancelarToolbar: Exit Function
        If tdbdSegNivel.Enabled = True And tdbdSegNivel.BoundText = "" Then Mensajes "Seleccione el 2do nivel de la lista": Call pSetFocus(tdbdSegNivel): CancelarToolbar: Exit Function
        If tdbdTerNivel.Enabled = True And tdbdTerNivel.BoundText = "" Then Mensajes "Seleccione el 3er nivel de la lista": Call pSetFocus(tdbdTerNivel): CancelarToolbar: Exit Function
    Else
        If tdbdPrimNivel.Visible = True And tdbdPrimNivel.BoundText = "" Then Mensajes "Seleccione el 1er nivel de la lista": Call pSetFocus(tdbdPrimNivel): CancelarToolbar: Exit Function
        If tdbdSegNivel.Visible = True And tdbdSegNivel.BoundText = "" Then Mensajes "Seleccione el 2do nivel de la lista": Call pSetFocus(tdbdSegNivel): CancelarToolbar: Exit Function
        If tdbdTerNivel.Visible = True And tdbdTerNivel.BoundText = "" Then Mensajes "Seleccione el 3er nivel de la lista": Call pSetFocus(tdbdTerNivel): CancelarToolbar: Exit Function
    End If
    
    If CE(tdbtCodigo.Text) = "" Then Mensajes "Falta Código.": Call pSetFocus(tdbtCodigo): CancelarToolbar: Exit Function
    If CE(tdbtNombre.Text) = "" Then Mensajes "Falta Descripción.": Call pSetFocus(tdbtNombre): CancelarToolbar: Exit Function

    ValidaDatos = True
End Function

Private Sub CancelarToolbar()
    Call BotonesVisibles(False, False, True, False, False, True, True)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grabar
' Description:       Procedimiento que graba el centro de costo
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Grabar()
    
    If ValidaDatos = False Then Exit Sub
    'If flgModificado Then
        If Mensajes("¿Desea Guardar los Datos?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        
        Dim clsMante As clsMantoTablas
        Set clsMante = New clsMantoTablas
'        Call AnimarProceso(False)
        ReDim arr(14) As Variant
        arr(0) = IIf(flgNuevo, "INSERTAR", "EDITAR")
        arr(1) = gsEmpresa
        arr(2) = gsAnio
        
        If arr(0) = "INSERTAR" Then
        Select Case tdbcNivel.BoundText
            'Case "P": arr(3) = tdbtCodigo.Text & "00000"
            Case "P": arr(3) = tdbtCodigo.Text
            'Case "S": arr(3) = tdbdPrimNivel.BoundText & tdbtCodigo.Text & "0000"
            Case "S": arr(3) = tdbdPrimNivel.BoundText & tdbtCodigo.Text
            'Case "T": arr(3) = tdbdPrimNivel.BoundText & tdbdSegNivel.BoundText & tdbtCodigo.Text & "00"
            Case "T": arr(3) = tdbdPrimNivel.BoundText & tdbdSegNivel.BoundText & tdbtCodigo.Text
            Case Else: arr(3) = tdbdPrimNivel.BoundText & tdbdSegNivel.BoundText & tdbdTerNivel.BoundText & tdbtCodigo.Text
        End Select
        Else
            arr(3) = tdbtCodigo.Text
        End If
        
        'arr(2) = tdbdPrimNivel.BoundText & tdbdSegNivel.BoundText & tdbtCodigo.Text
        arr(4) = "N"
        arr(5) = tdbtNombre.Text
        arr(6) = "A"
        arr(7) = ""
        arr(8) = gsUsuario
        arr(9) = ""
        arr(10) = ""
        arr(11) = "N"
        arr(12) = ""
        arr(13) = ""
        arr(14) = ""
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCNT_CENTRO_COSTO", arr(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar..."
            Call CancelarToolbar
            Set clsMante = Nothing
'            Call AnimarProceso(True)
            Exit Sub
        End If
        Set clsMante = Nothing
        Mensajes "Los datos se guardaron correctamente"
        Call BotonesVisibles(True, True, False, True, True, True, True)
        
        Call CargarListado
        ' LIMPIO EL FLAG
        flgModificado = False

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Eliminar
' Description:       Procedimiento que elimina el centro de costo seleccionado
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Eliminar()
    If Mensajes("Es UD. seguro de Eliminar el Registro :  " & Salto(2) & tdbgListado.Columns(1), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Dim clsMante As clsMantoTablas
        Set clsMante = New clsMantoTablas
'        Call AnimarProceso(False)
        ReDim arr(3) As Variant
        arr(0) = "ELIMINAR"
        arr(1) = gsEmpresa
        arr(2) = gsAnio
        arr(3) = tdbgListado.Columns(0)
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCNT_CENTRO_COSTO", arr(), True) = False Then
            Mensajes "El proceso no se ha realizado. Verificar..."
            Set clsMante = Nothing
'            Call AnimarProceso(True)
            Exit Sub
        End If
        Set clsMante = Nothing
        Call CargarListado
        Mensajes "Registro ha sido eliminado"
'        Call AnimarProceso(True)
    End If
End Sub


