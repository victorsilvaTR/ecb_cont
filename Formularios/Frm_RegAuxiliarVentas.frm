VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmManRegAuxiliarVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros Auxiliares"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   Icon            =   "Frm_RegAuxiliarVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   10905
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":220C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":25E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":29C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":2D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":3174
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6825
      Left            =   45
      TabIndex        =   31
      Top             =   405
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12039
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Consulta de Registros Auxiliares"
      TabPicture(0)   =   "Frm_RegAuxiliarVentas.frx":354E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Registros Auxiliares"
      TabPicture(1)   =   "Frm_RegAuxiliarVentas.frx":356A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrmTofo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame FrmTofo 
         Height          =   4695
         Left            =   135
         TabIndex        =   34
         Top             =   1575
         Width           =   10575
         Begin VB.TextBox tdbSerie 
            Height          =   315
            Left            =   1710
            TabIndex        =   16
            Top             =   600
            Width           =   1695
         End
         Begin VB.Frame fmrSoles 
            BorderStyle     =   0  'None
            Height          =   2895
            Left            =   8010
            TabIndex        =   73
            Top             =   135
            Width           =   2490
            Begin VB.Label lblConvBaseImp 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   675
               TabIndex        =   81
               Top             =   540
               Width           =   1680
            End
            Begin VB.Label lblConvBaseImpInaf 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   675
               TabIndex        =   80
               Top             =   900
               Width           =   1680
            End
            Begin VB.Label lblConvTotalDoc 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   675
               TabIndex        =   78
               Top             =   2025
               Width           =   1680
            End
            Begin VB.Label lblConvOtros 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   675
               TabIndex        =   77
               Top             =   1665
               Width           =   1680
            End
            Begin VB.Label lblTotalConv 
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   90
               TabIndex        =   76
               Top             =   2520
               Width           =   1065
            End
            Begin VB.Line Line2 
               BorderWidth     =   3
               X1              =   90
               X2              =   2295
               Y1              =   2430
               Y2              =   2430
            End
            Begin VB.Label lblConvTotal 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
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
               Left            =   450
               TabIndex        =   75
               Top             =   2520
               Width           =   1905
            End
            Begin VB.Label lblConvIgv 
               Alignment       =   1  'Right Justify
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   675
               TabIndex        =   74
               Top             =   1215
               Width           =   1680
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1590
            Left            =   180
            TabIndex        =   58
            Top             =   2970
            Width           =   10320
            Begin TDBText6Ctl.TDBText TdbApellidos 
               Height          =   315
               Left            =   5535
               TabIndex        =   28
               Top             =   1080
               Width           =   4515
               _Version        =   65536
               _ExtentX        =   7964
               _ExtentY        =   556
               Caption         =   "Frm_RegAuxiliarVentas.frx":3586
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_RegAuxiliarVentas.frx":35F2
               Key             =   "Frm_RegAuxiliarVentas.frx":3610
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
            Begin TDBText6Ctl.TDBText TdbNroDoc 
               Height          =   315
               Left            =   1530
               TabIndex        =   27
               Top             =   1080
               Width           =   2565
               _Version        =   65536
               _ExtentX        =   4524
               _ExtentY        =   556
               Caption         =   "Frm_RegAuxiliarVentas.frx":3654
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_RegAuxiliarVentas.frx":36C0
               Key             =   "Frm_RegAuxiliarVentas.frx":36DE
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
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   12
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
            Begin TrueOleDBList70.TDBCombo TdbTipoDoc 
               Height          =   300
               Left            =   1545
               TabIndex        =   26
               Tag             =   "enabled"
               Top             =   585
               Width           =   2580
               _ExtentX        =   4551
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
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
               Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
               Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
               Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
               Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
               Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
               Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
               _PropDict       =   $"Frm_RegAuxiliarVentas.frx":3730
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
            Begin TDBText6Ctl.TDBText tdbtEntidad 
               Height          =   300
               Left            =   1545
               TabIndex        =   24
               Top             =   180
               Width           =   960
               _Version        =   65536
               _ExtentX        =   1693
               _ExtentY        =   529
               Caption         =   "Frm_RegAuxiliarVentas.frx":37B7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_RegAuxiliarVentas.frx":3823
               Key             =   "Frm_RegAuxiliarVentas.frx":3841
               BackColor       =   16777152
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
            Begin TDBText6Ctl.TDBText TdbNombres 
               Height          =   300
               Left            =   2520
               TabIndex        =   25
               Top             =   180
               Width           =   7530
               _Version        =   65536
               _ExtentX        =   13282
               _ExtentY        =   529
               Caption         =   "Frm_RegAuxiliarVentas.frx":3883
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_RegAuxiliarVentas.frx":38EF
               Key             =   "Frm_RegAuxiliarVentas.frx":390D
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
            Begin VB.Label Label13 
               Caption         =   "Dirección"
               Height          =   255
               Left            =   4410
               TabIndex        =   61
               Top             =   1110
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "Nro. Documento"
               Height          =   255
               Left            =   180
               TabIndex        =   60
               Top             =   1080
               Width           =   1260
            End
            Begin VB.Label Label16 
               Caption         =   "Tipo de Doc."
               Height          =   255
               Left            =   180
               TabIndex        =   59
               Top             =   630
               Width           =   1215
            End
            Begin VB.Label lblnombre 
               Caption         =   "Cliente"
               Height          =   255
               Left            =   180
               TabIndex        =   82
               Top             =   225
               Width           =   735
            End
         End
         Begin TDBNumber6Ctl.TDBNumber TdbBaseImponible 
            Height          =   315
            Left            =   6240
            TabIndex        =   29
            Top             =   630
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":3951
            Caption         =   "Frm_RegAuxiliarVentas.frx":3971
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":39DD
            Keys            =   "Frm_RegAuxiliarVentas.frx":39FB
            Spin            =   "Frm_RegAuxiliarVentas.frx":3A53
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   8388608
            Format          =   "###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbIGV 
            Height          =   315
            Left            =   4815
            TabIndex        =   30
            Top             =   1395
            Width           =   525
            _Version        =   65536
            _ExtentX        =   926
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":3A7B
            Caption         =   "Frm_RegAuxiliarVentas.frx":3A9B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":3B07
            Keys            =   "Frm_RegAuxiliarVentas.frx":3B25
            Spin            =   "Frm_RegAuxiliarVentas.frx":3B7D
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   8388608
            Format          =   "###,###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   9999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBText6Ctl.TDBText TdbNumDocReg 
            Height          =   315
            Left            =   1710
            TabIndex        =   17
            Top             =   960
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "Frm_RegAuxiliarVentas.frx":3BA5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":3C11
            Key             =   "Frm_RegAuxiliarVentas.frx":3C2F
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
            MaxLength       =   12
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
         Begin TrueOleDBList70.TDBCombo cboTipoDocReg 
            Height          =   300
            Left            =   1710
            TabIndex        =   15
            Tag             =   "enabled"
            Top             =   270
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   7964
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
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
            AddItemSeparator=   ";"
            _PropDict       =   $"Frm_RegAuxiliarVentas.frx":3C81
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
         Begin TrueOleDBList70.TDBCombo tdbMoneda 
            Height          =   300
            Left            =   1710
            TabIndex        =   18
            Tag             =   "enabled"
            Top             =   1350
            Width           =   1695
            _ExtentX        =   2990
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
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
            _PropDict       =   $"Frm_RegAuxiliarVentas.frx":3D08
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
         Begin TDBNumber6Ctl.TDBNumber tdbBaseImpInafecto 
            Height          =   315
            Left            =   6255
            TabIndex        =   21
            Top             =   990
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":3D8F
            Caption         =   "Frm_RegAuxiliarVentas.frx":3DAF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":3E1B
            Keys            =   "Frm_RegAuxiliarVentas.frx":3E39
            Spin            =   "Frm_RegAuxiliarVentas.frx":3E73
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483634
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2293761
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber TdbTotal 
            Height          =   315
            Left            =   6255
            TabIndex        =   23
            Top             =   2130
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":3E9B
            Caption         =   "Frm_RegAuxiliarVentas.frx":3EBB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":3F27
            Keys            =   "Frm_RegAuxiliarVentas.frx":3F45
            Spin            =   "Frm_RegAuxiliarVentas.frx":3F7F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2293761
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbOtros 
            Height          =   315
            Left            =   6255
            TabIndex        =   22
            Top             =   1755
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":3FA7
            Caption         =   "Frm_RegAuxiliarVentas.frx":3FC7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":4033
            Keys            =   "Frm_RegAuxiliarVentas.frx":4051
            Spin            =   "Frm_RegAuxiliarVentas.frx":408B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2293761
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBDate6Ctl.TDBDate dtpFechaBus 
            Height          =   300
            Left            =   1710
            TabIndex        =   19
            Tag             =   "enabled"
            Top             =   1710
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Calendar        =   "Frm_RegAuxiliarVentas.frx":40B3
            Caption         =   "Frm_RegAuxiliarVentas.frx":41B5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":4219
            Keys            =   "Frm_RegAuxiliarVentas.frx":4237
            Spin            =   "Frm_RegAuxiliarVentas.frx":42A3
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
         Begin TDBNumber6Ctl.TDBNumber tdbTC 
            Height          =   315
            Left            =   6255
            TabIndex        =   20
            Top             =   270
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Calculator      =   "Frm_RegAuxiliarVentas.frx":42CB
            Caption         =   "Frm_RegAuxiliarVentas.frx":42EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":4357
            Keys            =   "Frm_RegAuxiliarVentas.frx":4375
            Spin            =   "Frm_RegAuxiliarVentas.frx":43AF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483634
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.000"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,##0.000"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   2293761
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSForms.CheckBox chkInafecto 
            Height          =   330
            Left            =   405
            TabIndex        =   79
            Top             =   2475
            Visible         =   0   'False
            Width           =   1050
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1852;582"
            Value           =   "0"
            Caption         =   "Inafecto"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblIGV 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   6255
            TabIndex        =   72
            Top             =   1350
            Width           =   1680
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   6930
            TabIndex        =   71
            Top             =   1755
            Width           =   1005
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Base Imp. Inafecto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4320
            TabIndex        =   64
            Top             =   1035
            Width           =   1620
         End
         Begin VB.Label lblTotalDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   5490
            TabIndex        =   70
            Top             =   2655
            Width           =   2445
         End
         Begin VB.Label lblTotal 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4320
            TabIndex        =   69
            Top             =   2700
            Width           =   1065
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   4365
            X2              =   7920
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Label Label19 
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4320
            TabIndex        =   65
            Top             =   1830
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Doc."
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
            Left            =   270
            TabIndex        =   57
            Top             =   285
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
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
            Left            =   270
            TabIndex        =   56
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nro. de Doc."
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
            Left            =   270
            TabIndex        =   55
            Top             =   960
            Width           =   1110
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Doc."
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
            Left            =   270
            TabIndex        =   54
            Top             =   1665
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "IGV (          ) %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4320
            TabIndex        =   53
            Top             =   1440
            Width           =   1305
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
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
            Left            =   270
            TabIndex        =   52
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label Label9 
            Caption         =   "Afecto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4320
            TabIndex        =   51
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Total Doc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4320
            TabIndex        =   50
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Valor T.C."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4320
            TabIndex        =   49
            Top             =   285
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   135
         ScaleHeight     =   885
         ScaleWidth      =   10515
         TabIndex        =   40
         Top             =   630
         Width           =   10545
         Begin VB.TextBox tdbtSaldoExt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   8820
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox tdbtMontoExt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   8820
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   90
            Width           =   1545
         End
         Begin VB.TextBox txtRegistro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            Height          =   330
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   990
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.TextBox tdbtMonto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   90
            Width           =   1545
         End
         Begin VB.TextBox tdbtSaldo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox tdbtMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   90
            Width           =   1545
         End
         Begin VB.TextBox tdbtFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox tdbtGlosa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   450
            Width           =   1590
         End
         Begin VB.TextBox tdbtNumVoucher 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   90
            Width           =   1590
         End
         Begin VB.Label lblSaldoUS 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Saldo US$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   7830
            TabIndex        =   67
            Top             =   495
            Width           =   915
         End
         Begin VB.Label lblTotalUS 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Total US$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   7830
            TabIndex        =   66
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Reg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   6
            Left            =   3330
            TabIndex        =   63
            Top             =   1080
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblTotalS 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Total S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   5400
            TabIndex        =   46
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   4
            Left            =   2970
            TabIndex        =   45
            Top             =   180
            Width           =   690
         End
         Begin VB.Label lblSaldoS 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Saldo S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   5355
            TabIndex        =   44
            Top             =   495
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   43
            Top             =   495
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Voucher"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   42
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   0
            Left            =   2970
            TabIndex        =   41
            Top             =   495
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5655
         Left            =   -74865
         TabIndex        =   32
         Top             =   405
         Width           =   10515
         Begin VB.OptionButton op_compras 
            Caption         =   "Registro de Auxiliar de Compras"
            Height          =   255
            Left            =   3660
            TabIndex        =   1
            Top             =   495
            Width           =   2655
         End
         Begin VB.OptionButton op_ventas 
            Caption         =   "Registro de Auxiliar de Ventas"
            Height          =   255
            Left            =   900
            TabIndex        =   0
            Top             =   495
            Width           =   2475
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgBoletas 
            Height          =   1650
            Left            =   150
            TabIndex        =   6
            Top             =   3870
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   2910
            _LayoutType     =   4
            _RowHeight      =   17
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Emp_cCodigo"
            Columns(0).DataField=   "Emp_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Pan_cAnio"
            Columns(1).DataField=   "Pan_cAnio"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Per_cPeriodo"
            Columns(2).DataField=   "Per_cPeriodo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Blta_cFlagLibro"
            Columns(3).DataField=   "Blta_cFlagLibro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tdo_cCodigo"
            Columns(4).DataField=   "Tdo_cCodigo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Reg. Nro."
            Columns(5).DataField=   "Blta_Correlativo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Serie"
            Columns(6).DataField=   "Asd_cSerieDoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Numero doc."
            Columns(7).DataField=   "Asd_cNumDoc"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Fecha doc."
            Columns(8).DataField=   "Asd_dFecha"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Base Imp. S/."
            Columns(9).DataField=   "Blta_nBaseImp"
            Columns(9).NumberFormat=   "Standard"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "IGV  S/."
            Columns(10).DataField=   "Blta_nIGV"
            Columns(10).NumberFormat=   "Standard"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Otros S/."
            Columns(11).DataField=   "Blta_nOtros"
            Columns(11).NumberFormat=   "Standard"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Total doc.  S/."
            Columns(12).DataField=   "Blta_nTotal"
            Columns(12).NumberFormat=   "Standard"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Moneda"
            Columns(13).DataField=   "Mon_cCodigo"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Blta_nTipoCambio"
            Columns(14).DataField=   "Blta_nTipoCambio"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Base Imp. US$"
            Columns(15).DataField=   "Blta_nBaseImpEXT"
            Columns(15).NumberFormat=   "Standard"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "IGV US$"
            Columns(16).DataField=   "Blta_nIGVEXT"
            Columns(16).NumberFormat=   "Standard"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Otros US$"
            Columns(17).DataField=   "Blta_nOtrosD"
            Columns(17).NumberFormat=   "Standard"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "Total doc. US$"
            Columns(18).DataField=   "Blta_nTotalEXT"
            Columns(18).NumberFormat=   "Standard"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "Tab_cTabla"
            Columns(19).DataField=   "Tab_cTabla"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(20)._VlistStyle=   0
            Columns(20)._MaxComboItems=   5
            Columns(20).Caption=   "Blta_num_doc"
            Columns(20).DataField=   "Blta_num_doc"
            Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(21)._VlistStyle=   0
            Columns(21)._MaxComboItems=   5
            Columns(21).Caption=   "Blta_cNombres"
            Columns(21).DataField=   "Blta_cNombres"
            Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(22)._VlistStyle=   0
            Columns(22)._MaxComboItems=   5
            Columns(22).Caption=   "Blta_cApellidos"
            Columns(22).DataField=   "Blta_cApellidos"
            Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(23)._VlistStyle=   0
            Columns(23)._MaxComboItems=   5
            Columns(23).Caption=   "Blta_cDNI"
            Columns(23).DataField=   "Blta_cDNI"
            Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(24)._VlistStyle=   0
            Columns(24)._MaxComboItems=   5
            Columns(24).Caption=   "T.C."
            Columns(24).DataField=   "Tca_nAuxiliar"
            Columns(24).NumberFormat=   "Standard"
            Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(25)._VlistStyle=   0
            Columns(25)._MaxComboItems=   5
            Columns(25).Caption=   "Mon_Codigo"
            Columns(25).DataField=   "Mon_Codigo"
            Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(26)._VlistStyle=   0
            Columns(26)._MaxComboItems=   5
            Columns(26).Caption=   "Blta_nBaseImpInaf"
            Columns(26).DataField=   "Blta_nBaseImpInaf"
            Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(27)._VlistStyle=   0
            Columns(27)._MaxComboItems=   5
            Columns(27).Caption=   "Blta_cInafecto"
            Columns(27).DataField=   "Blta_cInafecto"
            Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(28)._VlistStyle=   0
            Columns(28)._MaxComboItems=   5
            Columns(28).Caption=   "Blta_nBaseImpInafD"
            Columns(28).DataField=   "Blta_nBaseImpInafD"
            Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(29)._VlistStyle=   0
            Columns(29)._MaxComboItems=   5
            Columns(29).Caption=   "Total Nacional"
            Columns(29).DataField=   "Blta_nTotalDoc"
            Columns(29).NumberFormat=   "Standard"
            Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(30)._VlistStyle=   0
            Columns(30)._MaxComboItems=   5
            Columns(30).Caption=   "Total Extranjera"
            Columns(30).DataField=   "Blta_nTotalDocExt"
            Columns(30).NumberFormat=   "Standard"
            Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(31)._VlistStyle=   0
            Columns(31)._MaxComboItems=   5
            Columns(31).Caption=   "Entidad"
            Columns(31).DataField=   "Ent_cCodEntidad"
            Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   32
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=32"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(9)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(14)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(16)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(18)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(19)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(24)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(25)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(26)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(27)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(28)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(29)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(31)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(33)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(34)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(35)=   "Column(3).AllowFocus=0"
            Splits(0)._ColumnProps(36)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(37)=   "Column(4).Width=2302"
            Splits(0)._ColumnProps(38)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(4)._WidthInPix=2223"
            Splits(0)._ColumnProps(40)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(42)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(43)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(44)=   "Column(4).AllowFocus=0"
            Splits(0)._ColumnProps(45)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(46)=   "Column(5).Width=2170"
            Splits(0)._ColumnProps(47)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(5)._WidthInPix=2090"
            Splits(0)._ColumnProps(49)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(5)._ColStyle=529"
            Splits(0)._ColumnProps(51)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(52)=   "Column(6).Width=1693"
            Splits(0)._ColumnProps(53)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(6)._WidthInPix=1614"
            Splits(0)._ColumnProps(55)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(56)=   "Column(6)._ColStyle=529"
            Splits(0)._ColumnProps(57)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(58)=   "Column(7).Width=2672"
            Splits(0)._ColumnProps(59)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(7)._WidthInPix=2593"
            Splits(0)._ColumnProps(61)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(62)=   "Column(7)._ColStyle=529"
            Splits(0)._ColumnProps(63)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(64)=   "Column(8).Width=2593"
            Splits(0)._ColumnProps(65)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(66)=   "Column(8)._WidthInPix=2514"
            Splits(0)._ColumnProps(67)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(68)=   "Column(8)._ColStyle=529"
            Splits(0)._ColumnProps(69)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(70)=   "Column(9).Width=503"
            Splits(0)._ColumnProps(71)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(9)._WidthInPix=423"
            Splits(0)._ColumnProps(73)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(74)=   "Column(9).AllowSizing=0"
            Splits(0)._ColumnProps(75)=   "Column(9)._ColStyle=530"
            Splits(0)._ColumnProps(76)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(77)=   "Column(9).AllowFocus=0"
            Splits(0)._ColumnProps(78)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(79)=   "Column(10).Width=2302"
            Splits(0)._ColumnProps(80)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(81)=   "Column(10)._WidthInPix=2223"
            Splits(0)._ColumnProps(82)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(83)=   "Column(10).AllowSizing=0"
            Splits(0)._ColumnProps(84)=   "Column(10)._ColStyle=530"
            Splits(0)._ColumnProps(85)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(86)=   "Column(10).AllowFocus=0"
            Splits(0)._ColumnProps(87)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(88)=   "Column(11).Width=2328"
            Splits(0)._ColumnProps(89)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(90)=   "Column(11)._WidthInPix=2249"
            Splits(0)._ColumnProps(91)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(92)=   "Column(11).AllowSizing=0"
            Splits(0)._ColumnProps(93)=   "Column(11)._ColStyle=530"
            Splits(0)._ColumnProps(94)=   "Column(11).Visible=0"
            Splits(0)._ColumnProps(95)=   "Column(11).AllowFocus=0"
            Splits(0)._ColumnProps(96)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(97)=   "Column(12).Width=2302"
            Splits(0)._ColumnProps(98)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(99)=   "Column(12)._WidthInPix=2223"
            Splits(0)._ColumnProps(100)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(101)=   "Column(12).AllowSizing=0"
            Splits(0)._ColumnProps(102)=   "Column(12)._ColStyle=530"
            Splits(0)._ColumnProps(103)=   "Column(12).Visible=0"
            Splits(0)._ColumnProps(104)=   "Column(12).AllowFocus=0"
            Splits(0)._ColumnProps(105)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(106)=   "Column(13).Width=132"
            Splits(0)._ColumnProps(107)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(108)=   "Column(13)._WidthInPix=53"
            Splits(0)._ColumnProps(109)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(110)=   "Column(13).AllowSizing=0"
            Splits(0)._ColumnProps(111)=   "Column(13)._ColStyle=532"
            Splits(0)._ColumnProps(112)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(113)=   "Column(13).AllowFocus=0"
            Splits(0)._ColumnProps(114)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(115)=   "Column(14).Width=238"
            Splits(0)._ColumnProps(116)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(117)=   "Column(14)._WidthInPix=159"
            Splits(0)._ColumnProps(118)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(119)=   "Column(14).AllowSizing=0"
            Splits(0)._ColumnProps(120)=   "Column(14)._ColStyle=530"
            Splits(0)._ColumnProps(121)=   "Column(14).Visible=0"
            Splits(0)._ColumnProps(122)=   "Column(14).AllowFocus=0"
            Splits(0)._ColumnProps(123)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(124)=   "Column(15).Width=2778"
            Splits(0)._ColumnProps(125)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(126)=   "Column(15)._WidthInPix=2699"
            Splits(0)._ColumnProps(127)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(128)=   "Column(15).AllowSizing=0"
            Splits(0)._ColumnProps(129)=   "Column(15)._ColStyle=530"
            Splits(0)._ColumnProps(130)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(131)=   "Column(15).AllowFocus=0"
            Splits(0)._ColumnProps(132)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(133)=   "Column(16).Width=2540"
            Splits(0)._ColumnProps(134)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(135)=   "Column(16)._WidthInPix=2461"
            Splits(0)._ColumnProps(136)=   "Column(16)._EditAlways=0"
            Splits(0)._ColumnProps(137)=   "Column(16).AllowSizing=0"
            Splits(0)._ColumnProps(138)=   "Column(16)._ColStyle=530"
            Splits(0)._ColumnProps(139)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(140)=   "Column(16).AllowFocus=0"
            Splits(0)._ColumnProps(141)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(142)=   "Column(17).Width=2381"
            Splits(0)._ColumnProps(143)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(144)=   "Column(17)._WidthInPix=2302"
            Splits(0)._ColumnProps(145)=   "Column(17)._EditAlways=0"
            Splits(0)._ColumnProps(146)=   "Column(17).AllowSizing=0"
            Splits(0)._ColumnProps(147)=   "Column(17)._ColStyle=530"
            Splits(0)._ColumnProps(148)=   "Column(17).Visible=0"
            Splits(0)._ColumnProps(149)=   "Column(17).AllowFocus=0"
            Splits(0)._ColumnProps(150)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(151)=   "Column(18).Width=1402"
            Splits(0)._ColumnProps(152)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(153)=   "Column(18)._WidthInPix=1323"
            Splits(0)._ColumnProps(154)=   "Column(18)._EditAlways=0"
            Splits(0)._ColumnProps(155)=   "Column(18).AllowSizing=0"
            Splits(0)._ColumnProps(156)=   "Column(18)._ColStyle=530"
            Splits(0)._ColumnProps(157)=   "Column(18).Visible=0"
            Splits(0)._ColumnProps(158)=   "Column(18).AllowFocus=0"
            Splits(0)._ColumnProps(159)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(160)=   "Column(19).Width=212"
            Splits(0)._ColumnProps(161)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(162)=   "Column(19)._WidthInPix=132"
            Splits(0)._ColumnProps(163)=   "Column(19)._EditAlways=0"
            Splits(0)._ColumnProps(164)=   "Column(19).AllowSizing=0"
            Splits(0)._ColumnProps(165)=   "Column(19)._ColStyle=532"
            Splits(0)._ColumnProps(166)=   "Column(19).Visible=0"
            Splits(0)._ColumnProps(167)=   "Column(19).AllowFocus=0"
            Splits(0)._ColumnProps(168)=   "Column(19).Order=20"
            Splits(0)._ColumnProps(169)=   "Column(20).Width=2725"
            Splits(0)._ColumnProps(170)=   "Column(20).DividerColor=0"
            Splits(0)._ColumnProps(171)=   "Column(20)._WidthInPix=2646"
            Splits(0)._ColumnProps(172)=   "Column(20)._EditAlways=0"
            Splits(0)._ColumnProps(173)=   "Column(20).AllowSizing=0"
            Splits(0)._ColumnProps(174)=   "Column(20)._ColStyle=532"
            Splits(0)._ColumnProps(175)=   "Column(20).Visible=0"
            Splits(0)._ColumnProps(176)=   "Column(20).AllowFocus=0"
            Splits(0)._ColumnProps(177)=   "Column(20).Order=21"
            Splits(0)._ColumnProps(178)=   "Column(21).Width=106"
            Splits(0)._ColumnProps(179)=   "Column(21).DividerColor=0"
            Splits(0)._ColumnProps(180)=   "Column(21)._WidthInPix=26"
            Splits(0)._ColumnProps(181)=   "Column(21)._EditAlways=0"
            Splits(0)._ColumnProps(182)=   "Column(21).AllowSizing=0"
            Splits(0)._ColumnProps(183)=   "Column(21)._ColStyle=532"
            Splits(0)._ColumnProps(184)=   "Column(21).Visible=0"
            Splits(0)._ColumnProps(185)=   "Column(21).AllowFocus=0"
            Splits(0)._ColumnProps(186)=   "Column(21).Order=22"
            Splits(0)._ColumnProps(187)=   "Column(22).Width=2725"
            Splits(0)._ColumnProps(188)=   "Column(22).DividerColor=0"
            Splits(0)._ColumnProps(189)=   "Column(22)._WidthInPix=2646"
            Splits(0)._ColumnProps(190)=   "Column(22)._EditAlways=0"
            Splits(0)._ColumnProps(191)=   "Column(22).AllowSizing=0"
            Splits(0)._ColumnProps(192)=   "Column(22)._ColStyle=532"
            Splits(0)._ColumnProps(193)=   "Column(22).Visible=0"
            Splits(0)._ColumnProps(194)=   "Column(22).AllowFocus=0"
            Splits(0)._ColumnProps(195)=   "Column(22).Order=23"
            Splits(0)._ColumnProps(196)=   "Column(23).Width=2725"
            Splits(0)._ColumnProps(197)=   "Column(23).DividerColor=0"
            Splits(0)._ColumnProps(198)=   "Column(23)._WidthInPix=2646"
            Splits(0)._ColumnProps(199)=   "Column(23)._EditAlways=0"
            Splits(0)._ColumnProps(200)=   "Column(23).AllowSizing=0"
            Splits(0)._ColumnProps(201)=   "Column(23)._ColStyle=532"
            Splits(0)._ColumnProps(202)=   "Column(23).Visible=0"
            Splits(0)._ColumnProps(203)=   "Column(23).AllowFocus=0"
            Splits(0)._ColumnProps(204)=   "Column(23).Order=24"
            Splits(0)._ColumnProps(205)=   "Column(24).Width=1482"
            Splits(0)._ColumnProps(206)=   "Column(24).DividerColor=0"
            Splits(0)._ColumnProps(207)=   "Column(24)._WidthInPix=1402"
            Splits(0)._ColumnProps(208)=   "Column(24)._EditAlways=0"
            Splits(0)._ColumnProps(209)=   "Column(24).AllowSizing=0"
            Splits(0)._ColumnProps(210)=   "Column(24)._ColStyle=529"
            Splits(0)._ColumnProps(211)=   "Column(24).Visible=0"
            Splits(0)._ColumnProps(212)=   "Column(24).AllowFocus=0"
            Splits(0)._ColumnProps(213)=   "Column(24).Order=25"
            Splits(0)._ColumnProps(214)=   "Column(25).Width=2725"
            Splits(0)._ColumnProps(215)=   "Column(25).DividerColor=0"
            Splits(0)._ColumnProps(216)=   "Column(25)._WidthInPix=2646"
            Splits(0)._ColumnProps(217)=   "Column(25)._EditAlways=0"
            Splits(0)._ColumnProps(218)=   "Column(25).AllowSizing=0"
            Splits(0)._ColumnProps(219)=   "Column(25)._ColStyle=532"
            Splits(0)._ColumnProps(220)=   "Column(25).Visible=0"
            Splits(0)._ColumnProps(221)=   "Column(25).AllowFocus=0"
            Splits(0)._ColumnProps(222)=   "Column(25).Order=26"
            Splits(0)._ColumnProps(223)=   "Column(26).Width=2725"
            Splits(0)._ColumnProps(224)=   "Column(26).DividerColor=0"
            Splits(0)._ColumnProps(225)=   "Column(26)._WidthInPix=2646"
            Splits(0)._ColumnProps(226)=   "Column(26)._EditAlways=0"
            Splits(0)._ColumnProps(227)=   "Column(26).AllowSizing=0"
            Splits(0)._ColumnProps(228)=   "Column(26)._ColStyle=532"
            Splits(0)._ColumnProps(229)=   "Column(26).Visible=0"
            Splits(0)._ColumnProps(230)=   "Column(26).AllowFocus=0"
            Splits(0)._ColumnProps(231)=   "Column(26).Order=27"
            Splits(0)._ColumnProps(232)=   "Column(27).Width=2725"
            Splits(0)._ColumnProps(233)=   "Column(27).DividerColor=0"
            Splits(0)._ColumnProps(234)=   "Column(27)._WidthInPix=2646"
            Splits(0)._ColumnProps(235)=   "Column(27)._EditAlways=0"
            Splits(0)._ColumnProps(236)=   "Column(27).AllowSizing=0"
            Splits(0)._ColumnProps(237)=   "Column(27)._ColStyle=532"
            Splits(0)._ColumnProps(238)=   "Column(27).Visible=0"
            Splits(0)._ColumnProps(239)=   "Column(27).AllowFocus=0"
            Splits(0)._ColumnProps(240)=   "Column(27).Order=28"
            Splits(0)._ColumnProps(241)=   "Column(28).Width=2725"
            Splits(0)._ColumnProps(242)=   "Column(28).DividerColor=0"
            Splits(0)._ColumnProps(243)=   "Column(28)._WidthInPix=2646"
            Splits(0)._ColumnProps(244)=   "Column(28)._EditAlways=0"
            Splits(0)._ColumnProps(245)=   "Column(28).AllowSizing=0"
            Splits(0)._ColumnProps(246)=   "Column(28)._ColStyle=532"
            Splits(0)._ColumnProps(247)=   "Column(28).Visible=0"
            Splits(0)._ColumnProps(248)=   "Column(28).AllowFocus=0"
            Splits(0)._ColumnProps(249)=   "Column(28).Order=29"
            Splits(0)._ColumnProps(250)=   "Column(29).Width=4260"
            Splits(0)._ColumnProps(251)=   "Column(29).DividerColor=0"
            Splits(0)._ColumnProps(252)=   "Column(29)._WidthInPix=4180"
            Splits(0)._ColumnProps(253)=   "Column(29)._EditAlways=0"
            Splits(0)._ColumnProps(254)=   "Column(29)._ColStyle=530"
            Splits(0)._ColumnProps(255)=   "Column(29).Order=30"
            Splits(0)._ColumnProps(256)=   "Column(30).Width=1826"
            Splits(0)._ColumnProps(257)=   "Column(30).DividerColor=0"
            Splits(0)._ColumnProps(258)=   "Column(30)._WidthInPix=1746"
            Splits(0)._ColumnProps(259)=   "Column(30)._EditAlways=0"
            Splits(0)._ColumnProps(260)=   "Column(30)._ColStyle=530"
            Splits(0)._ColumnProps(261)=   "Column(30).Order=31"
            Splits(0)._ColumnProps(262)=   "Column(31).Width=2725"
            Splits(0)._ColumnProps(263)=   "Column(31).DividerColor=0"
            Splits(0)._ColumnProps(264)=   "Column(31)._WidthInPix=2646"
            Splits(0)._ColumnProps(265)=   "Column(31)._EditAlways=0"
            Splits(0)._ColumnProps(266)=   "Column(31)._ColStyle=532"
            Splits(0)._ColumnProps(267)=   "Column(31).Order=32"
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
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
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
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HF1EFEB&"
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=74,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=78,.parent=13,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=82,.parent=13,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=138,.parent=13,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=135,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=136,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=137,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
            _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=94,.parent=13,.alignment=1"
            _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=122,.parent=13,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=119,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=120,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=121,.parent=17"
            _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=118,.parent=13,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=115,.parent=14"
            _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=116,.parent=15"
            _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=117,.parent=17"
            _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=146,.parent=13,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=143,.parent=14"
            _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=144,.parent=15"
            _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=145,.parent=17"
            _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=43,.parent=14"
            _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=44,.parent=15"
            _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=45,.parent=17"
            _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=98,.parent=13"
            _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=95,.parent=14"
            _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=96,.parent=15"
            _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=97,.parent=17"
            _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=126,.parent=13"
            _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=123,.parent=14"
            _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=124,.parent=15"
            _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=125,.parent=17"
            _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=102,.parent=13"
            _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=99,.parent=14"
            _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=100,.parent=15"
            _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=101,.parent=17"
            _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=106,.parent=13"
            _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=103,.parent=14"
            _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=104,.parent=15"
            _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=105,.parent=17"
            _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=110,.parent=13"
            _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=107,.parent=14"
            _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=108,.parent=15"
            _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=109,.parent=17"
            _StyleDefs(134) =   "Splits(0).Columns(24).Style:id=114,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=111,.parent=14"
            _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=112,.parent=15"
            _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=113,.parent=17"
            _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=130,.parent=13"
            _StyleDefs(139) =   "Splits(0).Columns(25).HeadingStyle:id=127,.parent=14"
            _StyleDefs(140) =   "Splits(0).Columns(25).FooterStyle:id=128,.parent=15"
            _StyleDefs(141) =   "Splits(0).Columns(25).EditorStyle:id=129,.parent=17"
            _StyleDefs(142) =   "Splits(0).Columns(26).Style:id=134,.parent=13"
            _StyleDefs(143) =   "Splits(0).Columns(26).HeadingStyle:id=131,.parent=14"
            _StyleDefs(144) =   "Splits(0).Columns(26).FooterStyle:id=132,.parent=15"
            _StyleDefs(145) =   "Splits(0).Columns(26).EditorStyle:id=133,.parent=17"
            _StyleDefs(146) =   "Splits(0).Columns(27).Style:id=142,.parent=13"
            _StyleDefs(147) =   "Splits(0).Columns(27).HeadingStyle:id=139,.parent=14"
            _StyleDefs(148) =   "Splits(0).Columns(27).FooterStyle:id=140,.parent=15"
            _StyleDefs(149) =   "Splits(0).Columns(27).EditorStyle:id=141,.parent=17"
            _StyleDefs(150) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
            _StyleDefs(151) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
            _StyleDefs(152) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
            _StyleDefs(153) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
            _StyleDefs(154) =   "Splits(0).Columns(29).Style:id=154,.parent=13,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(155) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=14"
            _StyleDefs(156) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=15"
            _StyleDefs(157) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=17"
            _StyleDefs(158) =   "Splits(0).Columns(30).Style:id=158,.parent=13,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(159) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=14"
            _StyleDefs(160) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=15"
            _StyleDefs(161) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=17"
            _StyleDefs(162) =   "Splits(0).Columns(31).Style:id=162,.parent=13"
            _StyleDefs(163) =   "Splits(0).Columns(31).HeadingStyle:id=159,.parent=14"
            _StyleDefs(164) =   "Splits(0).Columns(31).FooterStyle:id=160,.parent=15"
            _StyleDefs(165) =   "Splits(0).Columns(31).EditorStyle:id=161,.parent=17"
            _StyleDefs(166) =   "Named:id=33:Normal"
            _StyleDefs(167) =   ":id=33,.parent=0"
            _StyleDefs(168) =   "Named:id=34:Heading"
            _StyleDefs(169) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(170) =   ":id=34,.wraptext=-1"
            _StyleDefs(171) =   "Named:id=35:Footing"
            _StyleDefs(172) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(173) =   "Named:id=36:Selected"
            _StyleDefs(174) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(175) =   "Named:id=37:Caption"
            _StyleDefs(176) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(177) =   "Named:id=38:HighlightRow"
            _StyleDefs(178) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(179) =   "Named:id=39:EvenRow"
            _StyleDefs(180) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(181) =   "Named:id=40:OddRow"
            _StyleDefs(182) =   ":id=40,.parent=33"
            _StyleDefs(183) =   "Named:id=41:RecordSelector"
            _StyleDefs(184) =   ":id=41,.parent=34"
            _StyleDefs(185) =   "Named:id=42:FilterBar"
            _StyleDefs(186) =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtCodigoBus 
            Height          =   315
            Left            =   4455
            TabIndex        =   5
            Top             =   3420
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   556
            Caption         =   "Frm_RegAuxiliarVentas.frx":43D7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":4443
            Key             =   "Frm_RegAuxiliarVentas.frx":4461
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
            MaxLength       =   12
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
         Begin TDBText6Ctl.TDBText tdbcorrelativo 
            Height          =   315
            Left            =   990
            TabIndex        =   4
            Top             =   3420
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "Frm_RegAuxiliarVentas.frx":44B3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Frm_RegAuxiliarVentas.frx":451F
            Key             =   "Frm_RegAuxiliarVentas.frx":453D
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
            MaxLength       =   12
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
         Begin TrueOleDBGrid70.TDBGrid tdbgAsientos 
            Height          =   1980
            Left            =   90
            TabIndex        =   3
            Top             =   1080
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   3493
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Interno"
            Columns(0).DataField=   "Ase_cNummov"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Año"
            Columns(1).DataField=   "PAN_CANIO"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fecha"
            Columns(2).DataField=   "Ase_dFecha"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Voucher"
            Columns(3).DataField=   "Ase_nVoucher"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Serie"
            Columns(4).DataField=   "Asd_cSerieDoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Numero"
            Columns(5).DataField=   "Asd_cNumDoc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nacional"
            Columns(6).DataField=   "Nac"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Extranjera"
            Columns(7).DataField=   "Ext"
            Columns(7).NumberFormat=   "Standard"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "CodMoneda"
            Columns(8).DataField=   "Ase_cTipoMoneda"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Moneda"
            Columns(9).DataField=   "Mon_cNombreLargo"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "TC"
            Columns(10).DataField=   "Ase_nTipoCambio"
            Columns(10).NumberFormat=   "External Editor"
            Columns(10).ExternalEditor=   "tdbtCambio"
            Columns(10).ExternalEditor.vt=   8
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Nombre Largo"
            Columns(11).DataField=   "Mon_cNombreLargo"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "TD."
            Columns(12).DataField=   "Asd_cTipoDoc"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Descripción"
            Columns(13).DataField=   "Tdo_cNombreLArgo"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   14
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowSizing=   -1  'True
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=14"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(9)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(14)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(15)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(16)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(18)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(19)=   "Column(2).Width=1614"
            Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=1535"
            Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=529"
            Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(25)=   "Column(3).Width=1905"
            Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=1826"
            Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=529"
            Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(31)=   "Column(4).Width=1191"
            Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1111"
            Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=529"
            Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(37)=   "Column(5).Width=2064"
            Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=1984"
            Splits(0)._ColumnProps(40)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=529"
            Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(43)=   "Column(6).Width=2275"
            Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2196"
            Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=530"
            Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(49)=   "Column(7).Width=2275"
            Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=2196"
            Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=530"
            Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(55)=   "Column(8).Width=1085"
            Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=1005"
            Splits(0)._ColumnProps(58)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
            Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=532"
            Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(62)=   "Column(8).AllowFocus=0"
            Splits(0)._ColumnProps(63)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(64)=   "Column(9).Width=2963"
            Splits(0)._ColumnProps(65)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(66)=   "Column(9)._WidthInPix=2884"
            Splits(0)._ColumnProps(67)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(68)=   "Column(9).AllowSizing=0"
            Splits(0)._ColumnProps(69)=   "Column(9)._ColStyle=532"
            Splits(0)._ColumnProps(70)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(71)=   "Column(9).AllowFocus=0"
            Splits(0)._ColumnProps(72)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(73)=   "Column(10).Width=926"
            Splits(0)._ColumnProps(74)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(10)._WidthInPix=847"
            Splits(0)._ColumnProps(76)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(77)=   "Column(10).AllowSizing=0"
            Splits(0)._ColumnProps(78)=   "Column(10)._ColStyle=530"
            Splits(0)._ColumnProps(79)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(80)=   "Column(10).AllowFocus=0"
            Splits(0)._ColumnProps(81)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(82)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(83)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(84)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(85)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(86)=   "Column(11).AllowSizing=0"
            Splits(0)._ColumnProps(87)=   "Column(11)._ColStyle=532"
            Splits(0)._ColumnProps(88)=   "Column(11).Visible=0"
            Splits(0)._ColumnProps(89)=   "Column(11).AllowFocus=0"
            Splits(0)._ColumnProps(90)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(91)=   "Column(12).Width=714"
            Splits(0)._ColumnProps(92)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(93)=   "Column(12)._WidthInPix=635"
            Splits(0)._ColumnProps(94)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(95)=   "Column(12)._ColStyle=529"
            Splits(0)._ColumnProps(96)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(97)=   "Column(13).Width=1746"
            Splits(0)._ColumnProps(98)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(99)=   "Column(13)._WidthInPix=1667"
            Splits(0)._ColumnProps(100)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(101)=   "Column(13)._ColStyle=532"
            Splits(0)._ColumnProps(102)=   "Column(13).Order=14"
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
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H80000008&"
            _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=Arial"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=Arial"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=91,.parent=1,.valignment=2,.bgcolor=&HF1EFEB&"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=100,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=92,.parent=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=93,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=94,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=96,.parent=6"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=95,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=97,.parent=8"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=98,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=99,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=101,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=102,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=106,.parent=91"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=92"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=93"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=95"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=110,.parent=91"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=92"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=93"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=95"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=114,.parent=91,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=92"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=93"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=95"
            _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=118,.parent=91,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=92"
            _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=93"
            _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=95"
            _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=122,.parent=91,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=119,.parent=92"
            _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=120,.parent=93"
            _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=121,.parent=95"
            _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=126,.parent=91,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=92"
            _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=93"
            _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=95"
            _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=130,.parent=91,.alignment=1,.bgcolor=&HCCFEFF&"
            _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=92"
            _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=93"
            _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=95"
            _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=134,.parent=91,.alignment=1,.bgcolor=&HFFDDCA&"
            _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=92"
            _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=93"
            _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=95"
            _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=138,.parent=91"
            _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=135,.parent=92"
            _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=136,.parent=93"
            _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=137,.parent=95"
            _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=142,.parent=91"
            _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=139,.parent=92"
            _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=140,.parent=93"
            _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=141,.parent=95"
            _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=146,.parent=91,.alignment=1"
            _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=143,.parent=92"
            _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=144,.parent=93"
            _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=145,.parent=95"
            _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=150,.parent=91"
            _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=147,.parent=92"
            _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=148,.parent=93"
            _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=149,.parent=95"
            _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=154,.parent=91,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=151,.parent=92"
            _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=152,.parent=93"
            _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=153,.parent=95"
            _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=158,.parent=91,.bgcolor=&HFFFFFF&"
            _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=155,.parent=92"
            _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=156,.parent=93"
            _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=157,.parent=95"
            _StyleDefs(93)  =   "Named:id=33:Normal"
            _StyleDefs(94)  =   ":id=33,.parent=0"
            _StyleDefs(95)  =   "Named:id=34:Heading"
            _StyleDefs(96)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(97)  =   ":id=34,.wraptext=-1"
            _StyleDefs(98)  =   "Named:id=35:Footing"
            _StyleDefs(99)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(100) =   "Named:id=36:Selected"
            _StyleDefs(101) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(102) =   "Named:id=37:Caption"
            _StyleDefs(103) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(104) =   "Named:id=38:HighlightRow"
            _StyleDefs(105) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(106) =   "Named:id=39:EvenRow"
            _StyleDefs(107) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(108) =   "Named:id=40:OddRow"
            _StyleDefs(109) =   ":id=40,.parent=33"
            _StyleDefs(110) =   "Named:id=41:RecordSelector"
            _StyleDefs(111) =   ":id=41,.parent=34"
            _StyleDefs(112) =   "Named:id=42:FilterBar"
            _StyleDefs(113) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList70.TDBCombo tdbcPeriodo 
            Height          =   300
            Left            =   7020
            TabIndex        =   2
            Tag             =   "enabled"
            Top             =   480
            Width           =   2850
            _ExtentX        =   5027
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
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
            _PropDict       =   $"Frm_RegAuxiliarVentas.frx":458F
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
         Begin VB.Label lblSaldoExt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "35,412.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8640
            TabIndex        =   68
            Top             =   3465
            Width           =   1650
         End
         Begin VB.Label lblMonSaldo 
            AutoSize        =   -1  'True
            Caption         =   "Saldo : S/."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            TabIndex        =   48
            Top             =   3465
            Width           =   1155
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "35,412.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8640
            TabIndex        =   47
            Top             =   3465
            Width           =   1650
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
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
            Index           =   0
            Left            =   7020
            TabIndex        =   39
            Top             =   180
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de la la Boleta"
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
            Index           =   7
            Left            =   90
            TabIndex        =   38
            Top             =   3105
            Width           =   1815
         End
         Begin VB.Label lblCabecera 
            AutoSize        =   -1  'True
            Caption         =   "Boletas de Ventas"
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
            Left            =   90
            TabIndex        =   37
            Top             =   765
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Registro"
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
            Index           =   3
            Left            =   90
            TabIndex        =   36
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Reg."
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
            Index           =   2
            Left            =   90
            TabIndex        =   35
            Top             =   3435
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Documento"
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
            Left            =   2970
            TabIndex        =   33
            Top             =   3480
            Width           =   1350
         End
      End
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
            Picture         =   "Frm_RegAuxiliarVentas.frx":4616
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4770
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":48CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":4F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":50E6
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
            Picture         =   "Frm_RegAuxiliarVentas.frx":5240
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":57DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":5D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":630E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":68A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":6E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":73DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":7976
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_RegAuxiliarVentas.frx":7F10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   83
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmManRegAuxiliarVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj_boletas As ADODB.Recordset
Dim rs_periodo  As ADODB.Recordset
'Dim rs_docs As ADODB.Recordset
Dim gtxtSQL As String
Dim obj As ClsFuncionesExecute

Dim flag_grabar As Boolean
Dim rs_boletas As ADODB.Recordset
Dim lrsAsientos As ADODB.Recordset
Dim BASE_EXT As Double
Dim IGV_EXT As Double
Dim OTROS As Double
Dim OTROSD As Double

Dim TOTAL_EXT As Double
Dim BASE_NAC As Double
Dim IGV_NAC As Double
Dim TOTAL_NAC As Double
Dim Tipo_Cambio As Double

Dim INAFECTO As Double
Dim INAFECTOD As Double

Dim lsLibroCom  As String
Dim lsLibroVen  As String
Dim nFilas  As Integer
Dim nFilasVoucher  As Integer
Dim gsGrupo As String
Dim nIGV As Double
Dim Indice As Integer
Dim lArrReg As New XArrayDB
Dim CargaInicial As Boolean

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Sub chkInafecto_Click()
    Hallar_montos
End Sub

Private Function ValidaFecha() As Boolean
    ValidaFecha = True
    If Format(dtpFechaBus, "yyyyMMdd") > Format(tdbtFecha, "yyyyMMdd") Then
        Mensajes "La fecha no debe ser mayor de la fecha del documento original ( " & CE(tdbtFecha.Text) & " )", vbInformation + vbOKOnly
        ValidaFecha = False 'lamando
        pSetFocus dtpFechaBus
    End If
End Function

Private Function validarDatos() As Boolean
    validarDatos = False
    Dim ValSaldo As Double
    Dim nTC  As Double
    Dim ValorSaldo As Double
    Dim ValorMonto As Double
        
    If tdbMoneda.BoundText = gsMonedaNac Then
        ValorSaldo = Round(CDbl(NE(tdbtSaldo.Text)), 2)
        ValorMonto = Round(NE(TdbTotal) + NE(tdbOtros) + NE(tdbBaseImpInafecto), 2)
    Else
        ValorSaldo = Round(CDbl(NE(tdbtSaldo.Text)), 2)
        ValorMonto = Round(CDbl(NE(lblConvTotal)), 2)
    End If
          
    If ValidaFecha = False Then
       Exit Function
    End If
    
    If Round(ValorSaldo, 2) - Round(ValorMonto, 2) <> 0 Then
       If MsgBox("El monto ingresado no cubre el saldo pendiente, Desea continuar ...", vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
            pSetFocus TdbTotal
            Exit Function
       End If
    End If
    
    If Round(ValorSaldo, 2) - Round(ValorMonto, 2) = 0 Then
        Mensajes "El monto ingresado cubre el saldo pendiente"
        'pSetFocus TdbTotal
        'Exit Function
    End If
    
    If op_compras.Value = True And Round(ValorMonto, 2) <= 0 Then
       Mensajes "El monto a ingresar debe ser mayor que cero", vbInformation + vbOKOnly
       pSetFocus TdbTotal
       Exit Function
    End If

    If Round(ValorMonto, 2) > Round(ValorSaldo, 2) Then
       Mensajes "El total del documento en " & gsNombreMonedaNac & " debe ser menor o igual al saldo mostrado", vbInformation + vbOKOnly
       pSetFocus TdbTotal
       Exit Function
    End If

    
    Dim Digitos  As Integer
    Digitos = BuscaTamanioDoc(TdbTipoDoc.BoundText)
    If Len(CE(TdbNroDoc.Text)) <> Digitos Then
        Mensajes "La cantidad de digitos del " & TdbTipoDoc.Text & " no es valido" & Salto(1) & "La cantidad de digitos debe ser : " & Digitos, vbInformation
        pSetFocus TdbNroDoc
        Exit Function
    End If
    
    If CE(TdbTipoDoc.Text) = "RUC" Then
       If Not fValidarNroRuc(TdbNroDoc.Text) Then
          Mensajes "El RUC no es valido", vbInformation
          pSetFocus TdbNroDoc
          Exit Function
       End If
    End If
    
    Dim UIT As Double
    Dim sUIT As Double
    UIT = Val(NE(BuscaValorEnConfigOP("027"))) / 2
    sUIT = UIT
    
    If tdbMoneda.BoundText = gsMonedaExt And tdbTC.Value > 0 Then
        UIT = UIT / tdbTC.Value
    End If
    
    If TdbTipoDoc.Enabled = True Then
            If (CE(TdbTipoDoc.BoundText) = "" Or _
               CE(TdbNroDoc) = "" Or _
               CE(TdbNombres) = "") And (NE(lblTotalDoc) > UIT) Then
               Mensajes "Se deben COMPLETAR los datos requeridos DEL DOCUMENTO" & Salto(2) & "El monto total paso de la media UIT = S/." & CE(sUIT), vbInformation + vbOKOnly
               pSetFocus TdbTipoDoc
            Exit Function
        End If
    End If

    'If TdbTipoCambio.Text = "" Then
    '   MsgBox "Verificar: Seleccione el Tipo de Cambio ", vbInformation, "":
    '   Exit Function
    'End If
    
    validarDatos = True
End Function

Private Sub GrabarDoc()
    DoEvents
    If validarDatos = False Then Exit Sub
    
    If Hallar_montos = True Then
        Call Grabar_editar
    Else
        Exit Sub
    End If
    
    Call CancelarDoc
    Call cargar_rs_boletas
    Call MuestraSaldo
    Me.SSTab1.TabEnabled(0) = True
    pSetFocus tdbgAsientos
    
    CargaInicial = False
End Sub

Private Function Hallar_montos() As Boolean
    'Evita que limpie el campo tipo de cambio cuando esta seleccionada la moneda soles
    If SSTab1.Tab = 0 Or CargaInicial = True And gintBiMoneda = 0 Then
        
        Hallar_montos = False
        Exit Function
    End If

    Dim IGV As Double
    
    TOTAL_NAC = 0
    IGV_NAC = 0
    BASE_NAC = 0
    INAFECTO = 0
    
    OTROS = 0
    OTROSD = 0
    
    TOTAL_EXT = 0
    IGV_EXT = 0
    BASE_EXT = 0
    INAFECTOD = 0
    
    If tdbMoneda.BoundText <> gsMonedaNac Then
         
       If NE(tdbTC.Value) = 0 Then
          Hallar_montos = False
          pSetFocus tdbTC
          Exit Function
       
       End If
    Else
       If tdbMoneda.BoundText = gsMonedaNac And gintBiMoneda <> 1 Then
          tdbTC.Value = 0
       'Solicita el ingreso del Tipo de Cambio
       ElseIf tdbMoneda.BoundText = gsMonedaNac And gintBiMoneda = 1 And CE(Me.tdbTC.Value) = 0 Then
          MsgBox "Debe ingresar el Tipo de Cambio, por tener activada la Opcion ByMoneda", vbExclamation, "Aviso.."
          Hallar_montos = False
          pSetFocus tdbTC
          Exit Function
       End If
    End If

    Tipo_Cambio = NE(tdbTC.Value)
    
    
    Hallar_montos = False
    
    IGV = NE(tdbIGV) / 100 ' Valor porcentual
    If Tipo_Cambio <= 0 And tdbMoneda.BoundText <> gsMonedaNac And SSTab1.Tab = 1 And CargaInicial = False Then
        
       'MsgBox "No se encuentra el tipo de cambio del dia", vbInformation
       tdbTC.Value = 1
       ' CargaFormularioTC
        Hallar_montos = False
        Exit Function
    End If
    
    If tdbMoneda.BoundText = gsMonedaNac Or tdbMoneda.Text = gsNombreMonedaNac Then
        TOTAL_NAC = NE(TdbTotal)
        OTROS = NE(tdbOtros)
        INAFECTO = tdbBaseImpInafecto
        BASE_NAC = NE(TOTAL_NAC) / (1 + NE(IGV))
        IGV_NAC = NE(BASE_NAC) * NE(IGV)
        
        TdbBaseImponible.Text = Format(NE(BASE_NAC), "###,###,##0.00")
        lblIGV.Caption = Format(IGV_NAC, "###,###,##0.00")
        
        If tdbMoneda.BoundText <> gsMonedaNac Then
            IGV_EXT = NE(IGV_NAC) / NE(Tipo_Cambio) 'NE(IGV)
            OTROSD = NE(tdbOtros) / NE(Tipo_Cambio)
            TOTAL_EXT = TOTAL_NAC / NE(Tipo_Cambio)
            BASE_EXT = NE(BASE_NAC) / NE(Tipo_Cambio)
            INAFECTOD = NE(INAFECTO) / NE(Tipo_Cambio)
        
        'Convierte los montos de Soles a Dolares
        ElseIf gintBiMoneda = 1 And tdbMoneda.BoundText = gsMonedaNac Then
            
            IGV_EXT = NE(IGV_NAC) / NE(Tipo_Cambio) 'NE(IGV)
            OTROSD = NE(tdbOtros) / NE(Tipo_Cambio)
            TOTAL_EXT = TOTAL_NAC / NE(Tipo_Cambio)
            BASE_EXT = NE(BASE_NAC) / NE(Tipo_Cambio)
            INAFECTOD = NE(INAFECTO) / NE(Tipo_Cambio)
            
            lblConvBaseImp = Format(NE(BASE_EXT), "###,###,##0.00")
            lblConvBaseImpInaf = Format(NE(INAFECTOD), "###,###,##0.00")
            lblConvIgv = Format(NE(IGV_EXT), "###,###,##0.00")
            lblConvOtros = Format(NE(OTROSD), "###,###,##0.00")
            
            lblConvTotalDoc = Format(NE(TOTAL_EXT), "###,###,##0.00")
            
            lblConvTotal = Format(BASE_EXT + INAFECTOD + IGV_EXT + OTROSD, "###,###,##0.00")
            
        End If
        
        lblTotalDoc.Caption = Format(BASE_NAC + INAFECTO + IGV_NAC + OTROS, "###,###,##0.00")
        
    Else
        
        TOTAL_EXT = NE(TdbTotal)
        OTROSD = NE(tdbOtros)
        INAFECTOD = tdbBaseImpInafecto
        BASE_EXT = NE(TOTAL_EXT) / (1 + NE(IGV))
        IGV_EXT = NE(BASE_EXT) * NE(IGV)
        
        TdbBaseImponible = Format(NE(BASE_EXT), "###,###,##0.00")
        lblIGV.Caption = Format(IGV_EXT, "###,###,##0.00")
        IGV_NAC = NE(IGV_EXT) * NE(Tipo_Cambio)
        INAFECTO = NE(INAFECTOD) * NE(Tipo_Cambio)
        OTROS = NE(OTROSD) * NE(Tipo_Cambio)
        TOTAL_NAC = NE(TOTAL_EXT) * NE(Tipo_Cambio)
        BASE_NAC = NE(BASE_EXT) * NE(Tipo_Cambio)
        lblConvBaseImp = Format(NE(BASE_NAC), "###,###,##0.00")
        lblConvBaseImpInaf = Format(NE(INAFECTO), "###,###,##0.00")
        lblConvIgv = Format(NE(IGV_NAC), "###,###,##0.00")
        lblConvOtros = Format(NE(OTROS), "###,###,##0.00")
        lblConvTotalDoc = Format(NE(TOTAL_NAC), "###,###,##0.00")
        lblConvTotal = Format(BASE_NAC + INAFECTO + IGV_NAC + OTROS, "###,###,##0.00")
        lblTotalDoc.Caption = Format(BASE_EXT + INAFECTOD + IGV_EXT + OTROSD, "###,###,##0.00")
        
    End If
    
    Hallar_montos = True
End Function

Private Sub CancelarDoc()
    activar_botones True, True, True, False, True
    Me.SSTab1.Tab = 0
    Me.SSTab1.TabEnabled(0) = True
    Me.SSTab1.TabEnabled(1) = False
    op_compras.Enabled = True
    op_ventas.Enabled = True
    tlbOpciones.Buttons(5).Enabled = True
    CargaInicial = True
    tdbtEntidad.Text = ""
End Sub

Private Sub Command2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape:
                If SSTab1.Tab = 0 Then
                    If MsgBox("Desea salir del registro de auxiliares", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        Unload Me
                    End If
                Else
                    If MsgBox("Desea cancelar el ingreso de este documento", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        Call CancelarDoc
                    End If
                End If
                
        Case vbKeyF2: If tlbOpciones.Buttons(1).Enabled Then ManNuevo
        'Case vbKeyF3: If tlbOpciones.Buttons(2).Enabled Then VerDatos
        Case vbKeyF4: GrabarDoc
        Case vbKeyF5: If tlbOpciones.Buttons(4).Enabled Then ManEliminar
        Case vbKeyF6: If tlbOpciones.Buttons(5).Enabled Then ManModificar
        Case vbKeyF7: If tlbOpciones.Buttons(5).Enabled Then FrmRepAuxiliarVentas.Show
    End Select
    
End Sub

'Private Sub Grabar()
'    If cmd_grabar.Enabled = True And SSTab1.Tab = 1 Then
'        Call cmd_grabar_Click
'    End If
'End Sub

Private Sub pCargaCfgLibro()
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    sqlver = "SELECT * From CNT_CONFIG_LIBROS WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio='" & gsAnio & "'"
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
       lsLibroCom = CE(rsArreglo("Cfl_cCompras"))
       lsLibroVen = CE(rsArreglo("Cfl_cVentas"))
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Sub
Private Sub CargaTabla()
    Dim RegFiltro As String
    If op_compras.Value = True Then
        RegFiltro = lsLibroCom
    Else
        RegFiltro = lsLibroVen
    End If

    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Set lrsAsientos = New ADODB.Recordset
    Set lrsAsientos.DataSource = Nothing
    tdbgAsientos.DataSource = lrsAsientos
    tdbgAsientos.ReBind
    
    nFilasVoucher = 0
    sqlSp = "spCn_ConsultaAsientos 'SEL_ALLCAB_REGAUX', '', '" & gsEmpresa & _
               "', '" & gsAnio & "', '" & tdbcPeriodo.BoundText & "', '" & RegFiltro & "', '', '', '', '', ''"
    
    arrDatos = Array(sqlSp)
    Set lrsAsientos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsAsientos Is Nothing Then
       
       'lrsAsientos.Sort = "Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher"
       tdbgAsientos.DataSource = lrsAsientos
       tdbgAsientos.ReBind
       
        nFilasVoucher = lrsAsientos.RecordCount
    End If
    
    
    
    Set clDatos = Nothing
    Set lrsAsientos = Nothing
    
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    CargaInicial = True
    Dim periodo  As String
    Call Centrar_form(Me)
    
    Set obj = New ClsFuncionesExecute

    pCargaCfgLibro
    Call CargaTabla

    cboTipoDocReg.Text = ""
    SSTab1.Tab = 0
    op_ventas.Value = True
    
    SSTab1.TabEnabled(1) = False
    Call Carga_combos
    Call Habilitar_frame_nombres(True, True)
    
    
    activar_botones True, True, True, False
    
    tdbcPeriodo.BoundText = gsPeriodo
    pFiltrarDatos
    
    op_compras.Enabled = True
    op_ventas.Enabled = True
    
    Me.tdbgBoletas.HighlightRowStyle = "HighlightRow"
    Me.tdbgAsientos.HighlightRowStyle = "HighlightRow"
    
    SeteaBarraHerramientas Me.tlbOpciones, gsGrupo
    
    nIGV = fIgv
    tdbIGV = nIGV
    
    dtpFechaBus.Value = "01/01/" & gsAnio
    dtpFechaBus.MinDate = "01/01/" & gsAnio
    dtpFechaBus.MaxDate = "31/12/" & gsAnio
    
    If gsPeriodo = "00" Then
        periodo = "01"
    ElseIf gsPeriodo > "12" Then
        periodo = "12"
    Else
        periodo = gsPeriodo
    End If

    dtpFechaBus.Value = "01/" & periodo & "/" & gsAnio
     
    If ExistenDatosOp("028") = False Then
       Mensajes "Faltan datos en configuración de operaciones, en el rubro" & Salto(2) & "028: CODIGO DE DOCUMENTO PARA REGISTROS AUXILIARES", vbOKOnly + vbInformation
       
       DesactivaBarraHerramientas False
    Else
        SeteaBarraHerramientas Me.tlbOpciones, gsGrupo
    End If

    cboTipoDocReg.Locked = True
End Sub

Private Sub DesactivaBarraHerramientas(Valor As Boolean)
    tlbOpciones.Buttons(1).Enabled = Valor
    tlbOpciones.Buttons(2).Enabled = Valor
    tlbOpciones.Buttons(3).Enabled = Valor
    tlbOpciones.Buttons(4).Enabled = Valor
    tlbOpciones.Buttons(5).Enabled = Valor
    tlbOpciones.Buttons(6).Enabled = Valor
End Sub


Private Sub cargar_rs_boletas()
    
    Dim cadena As String
    If op_compras.Value = True Then
        cadena = "C"
    Else
        cadena = "V"
    End If

    gtxtSQL = "spCNT_REG_BOLETAS 'BUSCARTODOS','" & gsEmpresa & "','" & gsAnio & "',NULL,'" & cadena & "'," & _
              "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
              "NULL,NULL,NULL,NULL,NULL,'" & CE(tdbgAsientos.Columns(0)) & "','" & CE(tdbgAsientos.Columns(3)) & "','" & CE(tdbgAsientos.Columns(4)) & "','" & CE(tdbgAsientos.Columns(5)) & "',NULL,NULL,NULL,''"

    
    LlenarArreglo lArrReg, gtxtSQL
    
    Set tdbgBoletas.Array = lArrReg
    Set rs_boletas = New ADODB.Recordset
    If rs_boletas.State = adStateOpen Then rs_boletas.Close
    rs_boletas.CursorLocation = adUseClient
    Set rs_boletas = obj.fRetornaRS(gtxtSQL)
    Set tdbgBoletas.DataSource = rs_boletas
    
    nFilas = rs_boletas.RecordCount
End Sub

Private Sub Carga_combos()
    '--------------------------------------------------------------------------'
    gtxtSQL = "SELECT tab_ccodigo , tab_cdescripcampo FROM TABLA WHERE Emp_cCodigo = '" & gsEmpresa & "' and Tab_cTabla='003' "
    LlenarComboAddItem TdbTipoDoc, gtxtSQL, True
    '--------------------------------------------------------------------------'
    gtxtSQL = "SELECT Per_cPeriodo, Per_cDescripPeriodo  FROM cnt_periodo where Emp_cCodigo='" & gsEmpresa & "' and " & _
              "Pan_cAnio='" & gsAnio & "'"
    LlenarComboAddItem tdbcPeriodo, gtxtSQL
    '--------------------------------------------------------------------------'
    gtxtSQL = "select Mon_cCodigo, Mon_cNombreLargo from CNT_TIPO_MONEDA where Emp_cCodigo = '" & gsEmpresa & "' and ( mon_cmnac='1' or mon_cmext='1'  )"
    LlenarComboAddItem tdbMoneda, gtxtSQL
    '--------------------------------------------------------------------------'
    gtxtSQL = "select DISTINCT OP.COD_CVALORPARAM , TD.TDO_CNOMBRELARGO " & _
              "from CND_CONFIG_OPERA OP LEFT JOIN cNT_TIPODOC TD ON OP.EMP_CCODIGO = TD.EMP_CCODIGO AND OP.COD_CVALORPARAM = TD.TDO_CCODIGO " & _
              "where OP.cop_ccodigo='028' AND  OP.EMP_CCODIGO='" & gsEmpresa & "' AND OP.PAN_CANIO='" & gsAnio & "'"
    LlenarComboAddItem cboTipoDocReg, gtxtSQL
End Sub


Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTab1
            .Width = Me.Width - .Left + 15 - 200
            .Height = Me.Height - .Top + 15 - 500
            
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 200
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 200
        End With
       
        With tdbgAsientos
            .Width = Frame1.Width - .Left - 200
        End With
        
        With tdbgBoletas
            .Width = Frame1.Width - .Left - 200
            .Height = Frame1.Height - .Top - 200
        End With
                
        Picture1.Width = SSTab1.Width - 300
        
        FrmTofo.Width = Picture1.Width
        FrmTofo.Height = SSTab1.Height - FrmTofo.Top - 200
        
        Frame2.Width = FrmTofo.Width - 300
        Frame2.Height = FrmTofo.Height - Frame2.Top - 150
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub



Private Sub op_compras_Click()
    lblCabecera.Caption = "Documento del Registro de Compras"
    tdbOtros.Visible = True
    tdbOtros.Enabled = True
    SSTab1.TabCaption(1) = " Mantenimiento de Registros Auxiliares de Compras "
    
    pFiltrarDatos
    Call CargaTabla
End Sub

Private Sub op_ventas_Click()
    lblCabecera.Caption = "Documento del Registro de Ventas"
    tdbOtros.Visible = True
    tdbOtros.Enabled = True
    SSTab1.TabCaption(1) = " Mantenimiento de Registros Auxiliares de Ventas "
    
    pFiltrarDatos
    Call CargaTabla
End Sub



Private Sub TdbApellidos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GrabarDoc
    End If
End Sub
'Private Sub TDBCombo2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then pSendKeys "{TAB}"
'End Sub

Private Sub tdbBaseImpInafecto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ControlAbs tdbBaseImpInafecto
        CargaInicial = False
        'Hallar_montos
        
        If CalculaMonto(NE(TdbTotal.Value), NE(tdbOtros.Value), NE(tdbBaseImpInafecto.Value)) = False Then
            tdbBaseImpInafecto.Value = 0
            pSetFocus tdbBaseImpInafecto
            Hallar_montos
            KeyCode = 0
        Else
            pSetFocus tdbOtros
            KeyCode = 0
            
        End If
        
    End If
End Sub

Private Sub tdbBaseImpInafecto_LostFocus()
    Hallar_montos
    ActivaUIT
    
End Sub

Private Sub tdbcorrelativo_Change()
    pFiltrarDatos
End Sub
Private Sub tdbcorrelativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Function SumarTotales(Columna As Integer) As Double
    Dim i As Integer
    Dim Suma As Double
    Suma = 0
    
    For i = 0 To nFilas - 1
        Suma = Suma + NE(lArrReg(i, Columna))
    Next i
    SumarTotales = Suma
End Function

Private Function SumarSaldoPendiente(Columna As Integer) As Double
    Dim i As Integer
    Dim Suma As Double
    Dim Correlativo As String
    Suma = 0
    
    Correlativo = CE(tdbgBoletas.Columns(5).Value)
    
    For i = 0 To nFilas - 1
        If CE(lArrReg(i, 5)) = Correlativo Then
            Suma = Suma + NE(lArrReg(i, Columna))
        End If
    Next i
    SumarSaldoPendiente = Suma
End Function

Private Sub MuestraSaldoPorMoneda()
    lblTotalS.Visible = True
    lblSaldoS.Visible = True
    tdbtMonto.Visible = True
    tdbtSaldo.Visible = True
End Sub


Private Sub MuestraSaldo()
    Dim Saldo As Double
    
        lblMonSaldo.Caption = "Saldo : S/."
        lblSaldo.Visible = True
        lblSaldoExt.Visible = False
     
'    lblSaldo.Caption = Format(NE(tdbgAsientos.Columns(6).Value) - SumarTotales(11) - SumarTotales(31) - SumarTotales(30), "###,###,##0.00")
'    tdbtSaldo.Text = Format(NE(lblSaldo.Caption) + SumarSaldoPendiente(11) + SumarSaldoPendiente(31) + SumarSaldoPendiente(30), "###,###,##0.00")

    lblSaldo.Caption = Round(NE(tdbgAsientos.Columns(6).Value) - SumarTotales(11) - SumarTotales(31) - SumarTotales(30), 2)
    tdbtSaldo.Text = Round(NE(lblSaldo.Caption) + SumarSaldoPendiente(11) + SumarSaldoPendiente(31) + SumarSaldoPendiente(30), 2)
    
    If NE(tdbgAsientos.Columns(6).Value) = 0 Then tdbtSaldo.Text = "0.00"
End Sub

Private Sub tdbcPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSetFocus tdbgAsientos
    End If
End Sub

Private Sub tdbgAsientos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Me.tdbtFecha.Text = tdbgAsientos.Columns(2)
    Me.tdbtNumVoucher.Text = tdbgAsientos.Columns(3)
    Me.tdbtGlosa.Text = "  " & CE(tdbgAsientos.Columns(4)) & " - " & CE(tdbgAsientos.Columns(5))
    Me.tdbtMonto.Text = tdbgAsientos.Columns(6)
    Me.tdbtMontoExt.Text = tdbgAsientos.Columns(7)
    Me.tdbtMoneda.Text = tdbgAsientos.Columns(11)
    tdbMoneda.BoundText = tdbgAsientos.Columns(8)
    
    
    OcultaColumnasMoneda
    
    cargar_rs_boletas
    
    MuestraSaldo
End Sub

Private Sub OcultaColumnasMoneda()
    lblTotalS.Visible = True
    lblSaldoS.Visible = True
    tdbtMonto.Visible = True
    tdbtSaldo.Visible = True
End Sub


Private Sub tdbIGV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSetFocus tdbOtros
    End If
End Sub

Private Sub tdbMoneda_ItemChange()

    If tdbMoneda.BoundText = gsMonedaNac Then
        ' TdbTipoCambio.BoundText = "SCV"
        lblTotal.Caption = "Total " & gsMonedaNacAbrev
        fmrSoles.Visible = False
        
        lblTotalUS.Visible = False
        lblSaldoUS.Visible = False
        tdbtMontoExt.Visible = False
        tdbtSaldoExt.Visible = False
        
        
        
        lblConvBaseImp.Caption = "0.00"
        lblConvBaseImpInaf.Caption = "0.00"
        lblConvIgv.Caption = "0.00"
        lblConvOtros.Caption = "0.00"
        lblConvTotal.Caption = "0.00"
        lblConvTotalDoc.Caption = "0.00"
        
        'Cambia el titulo de los Label
        If gintBiMoneda = 1 Then
            lblTotal.Caption = "Total " & gsMonedaNacAbrev
            lblTotalConv.Caption = "Total " & gsMonedaExtAbrev
            fmrSoles.Visible = True
            lblTotalUS.Visible = False
            lblSaldoUS.Visible = False
            tdbtMontoExt.Visible = False
            tdbtSaldoExt.Visible = False
        Else
            tdbTC.Value = 0
        End If
        
    Else
        'If TdbTipoCambio.BoundText = "SCV" Then TdbTipoCambio.BoundText = "VEP"
        lblTotal.Caption = "Total " & gsMonedaExtAbrev
        lblTotalConv.Caption = "Total " & gsMonedaNacAbrev
        

        
        fmrSoles.Visible = True
        lblTotalUS.Visible = False
        lblSaldoUS.Visible = False
        tdbtMontoExt.Visible = False
        tdbtSaldoExt.Visible = False
        
    End If
    
    
    
End Sub

Private Sub tdbMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CargaInicial = False
    pSendKeys "{TAB}"
End If
End Sub

'Private Sub tdbnAño_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then pSendKeys "{TAB}"
'End Sub

Private Sub TdbNombres_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub TdbNroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub TdbNumDocReg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TdbNumDocReg.Text = Right("00000000" & CE(TdbNumDocReg.Text), 8)
        TdbNumDocReg.SelStart = Len(CE(TdbNumDocReg.Text))
    End If
End Sub

Private Sub TdbNumDocReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbcPeriodo_Change()
pFiltrarDatos
End Sub

Private Sub tdbcPeriodo_ItemChange()
    Call CargaTabla
End Sub

Private Sub tdbcPeriodoKeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub
Private Sub cboTipoDocReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbOtros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ControlAbs tdbOtros
        CargaInicial = False
        If CalculaMonto(NE(TdbTotal.Value), NE(tdbOtros.Value), NE(tdbBaseImpInafecto.Value)) = False Then
            tdbOtros.Value = 0
            pSetFocus tdbOtros
            KeyCode = 0
            Hallar_montos
        Else
            KeyCode = 0
            pSetFocus TdbTotal
        End If
        
    End If
End Sub

Private Sub tdbOtros_LostFocus()
    Call ActivaUIT
    Call Hallar_montos
End Sub

Private Sub tdbSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim Serie As String
        Serie = CE(tdbSerie.Text)
        If Len(Serie) < 3 And Serie <> "" Then Serie = Right("000" & Serie, 3)
    
        tdbSerie.Text = Serie
        tdbSerie.SelStart = Len(CE(tdbSerie.Text))
    End If
End Sub

Private Sub tdbSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbTC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If tdbMoneda.BoundText <> gsMonedaNac And tdbTC.Value = 0 Then
           Mensajes "Ingrese un tipo de cambio", vbOKOnly + vbInformation
           pSetFocus tdbTC
        Else
           Hallar_montos
           pSetFocus tdbBaseImpInafecto
        End If
    
    End If
End Sub

Private Sub tdbTC_LostFocus()
    tdbTC.Value = NE(tdbTC.Value)
    
    If tdbMoneda.BoundText = gsMonedaNac And gintBiMoneda = 0 Then
       tdbTC.Value = 0
    End If
End Sub

Private Sub tdbtCodigoBus_Change()
    Call pFiltrarDatos
End Sub

Sub pFiltrarDatos()
Dim cadena As String
Dim filtros(2) As String
Dim i As Integer
    If rs_boletas Is Nothing Then Exit Sub
    cadena = ""
    If Trim(tdbcorrelativo) <> "" Then filtros(0) = "Blta_Correlativo like '" & tdbcorrelativo.Text & "*'"
    If Trim(tdbtCodigoBus) <> "" Then filtros(1) = "Asd_cNumDoc like '" & tdbtCodigoBus.Text & "*'"
    If tdbcPeriodo.BoundText <> "" Then filtros(2) = "Per_cPeriodo like '" & tdbcPeriodo.BoundText & "*'"
    For i = 0 To 2
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        
        End If
    Next
    If Trim(cadena) <> "" Then
        rs_boletas.Filter = cadena
    Else
        rs_boletas.Filter = 0
     End If
Set tdbgBoletas.DataSource = rs_boletas
End Sub

'Private Sub TdbTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        CargaInicial = False
'
'        If tdbOtros.Enabled = True Then
'            KeyCode = 0
'            pSetFocus tdbBaseImpInafecto
'            KeyCode = 0
'            DoEvents
'        Else
'            pSetFocus TdbTotal
'            KeyCode = 0
'        End If
'
'        Hallar_montos
'    End If
'End Sub

Private Sub TdbTipoDoc_ItemChange()
    If TdbTipoDoc.BoundText = "" Then TdbNroDoc.Text = ""
    TdbNroDoc.MaxLength = BuscaTamanioDoc(TdbTipoDoc.BoundText)
End Sub

Private Sub TdbTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If TdbTipoDoc.BoundText = "" Then
            Call GrabarDoc
            KeyCode = 0
        Else
            pSetFocus TdbNroDoc
        End If
        
    End If

End Sub

Private Function BuscaValorEnConfigOP(Codigo As String) As String
    Dim sqlver  As String, valorDato As String
    sqlver = "SELECT Cod_cValorParam FROM CND_CONFIG_OPERA " & _
             "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and  " & _
             "cop_ccodigo='" & Codigo & "' "
    valorDato = ExtraeDescripcion(sqlver)
    If valorDato <> "" Then
        BuscaValorEnConfigOP = CE(valorDato)
    Else
        BuscaValorEnConfigOP = ""
    End If
End Function

'Private Sub cboTipoDoc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then pSendKeys "{TAB}"
'End Sub
Private Sub Grabar_editar()
ReDim arr(38) As Variant
arr(0) = IIf(flag_grabar = True, "'INSERTAR'", "'EDITAR'")
arr(1) = "'" & gsEmpresa & "'"
arr(2) = "'" & gsAnio & "'"
arr(3) = "'" & CE(tdbcPeriodo.BoundText) & "'"
arr(4) = IIf(op_compras.Value = True, "'C'", "'V'")
arr(5) = "'" & CE(cboTipoDocReg.BoundText) & "'"                           ' CNT_TIPODOC >> Codigo de Boleta = 03
arr(6) = "'" & CE(txtRegistro.Text) & " '"   'REGISTRO
arr(7) = "'" & CE(tdbSerie) & "'"
arr(8) = "'" & CE(TdbNumDocReg) & "'"
arr(9) = "'" & CE(dtpFechaBus) & "'"


arr(10) = NE(BASE_NAC)
arr(11) = NE(IGV_NAC)

arr(12) = NE(TOTAL_NAC)



arr(13) = "'" & tdbMoneda.BoundText & "'"
arr(14) = "''" ' & IIf(TdbTipoCambio.BoundText <> "", TdbTipoCambio.BoundText, "MAN") & "'"
arr(15) = NE(Tipo_Cambio)
arr(16) = NE(BASE_EXT)
arr(17) = NE(IGV_EXT)
arr(18) = NE(TOTAL_EXT)
arr(19) = "'" & CE(TdbTipoDoc.BoundText) & "'" 'TdbTipoDoc
arr(20) = "'" & TdbNroDoc.Text & "'"
arr(21) = "'" & TdbNombres.Text & "'"
arr(22) = "'" & TdbApellidos.Text & "'"
arr(23) = "'" & TdbNroDoc.Text & "'"
arr(25) = "'" & gsUsuario & "'"
arr(27) = "'" & gsUsuario & "'"
arr(30) = "'" & CE(tdbgAsientos.Columns(0)) & "'"
arr(31) = "'" & CE(tdbgAsientos.Columns(3)) & "'"
arr(32) = "'" & CE(tdbgAsientos.Columns(4)) & "'"
arr(33) = "'" & CE(tdbgAsientos.Columns(5)) & "'"

arr(34) = NE(INAFECTO)
arr(35) = NE(OTROS)
arr(36) = NE(OTROSD)
arr(37) = "'" & NE(chkInafecto.Value) & "'"
arr(38) = NE(INAFECTOD)


obj.Mant_Tablas arr, "SpCNT_REG_BOLETAS", 38


End Sub
Private Sub Eliminar()
    ReDim arr(38) As Variant
    arr(0) = "'ELIMINAR'"
    arr(1) = "'" & gsEmpresa & "'"
    arr(2) = "'" & gsAnio & "'"
    arr(3) = "'" & tdbgBoletas.Columns(2) & "'"
    arr(4) = "'" & tdbgBoletas.Columns(3) & "'"
    arr(5) = "'" & tdbgBoletas.Columns(4) & "'"
    arr(6) = "'" & tdbgBoletas.Columns(5) & "'"
    arr(25) = "'" & gsUsuario & "'"
    arr(30) = "'" & CE(tdbgAsientos.Columns(0)) & "'"
    arr(31) = "'" & CE(tdbgAsientos.Columns(3)) & "'"
    arr(32) = "'" & CE(tdbgAsientos.Columns(4)) & "'"
    arr(33) = "'" & CE(tdbgAsientos.Columns(5)) & "'"
    arr(34) = NE(INAFECTO)
    arr(35) = NE(tdbOtros)
    arr(36) = NE(tdbOtros)
    arr(37) = "'" & NE(chkInafecto.Value) & "'"
    arr(38) = NE(INAFECTOD)
    
    obj.Mant_Tablas arr, "SpCNT_REG_BOLETAS", 38
End Sub

Private Sub Habilitar_frame_nombres(Valor As Boolean, Limpiar As Boolean)
    Frame2.Enabled = Valor
    TdbTipoDoc.Enabled = Valor
    TdbNroDoc.Enabled = Valor
    TdbNombres.Enabled = Valor
    TdbApellidos.Enabled = Valor

    If Limpiar = True Then
        TdbTipoDoc.BoundText = ""
        TdbNroDoc = ""
        TdbNombres = ""
        TdbApellidos = ""
    End If
End Sub

Private Sub TdbTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ControlAbs TdbTotal
        CargaInicial = False
        If CalculaMonto(NE(TdbTotal.Value), NE(tdbOtros.Value), NE(tdbBaseImpInafecto.Value)) = False Then
            TdbTotal.Value = 0
            pSetFocus TdbTotal
            KeyCode = 0
            Hallar_montos
        Else
            KeyCode = 0
            pSetFocus tdbtEntidad
        
        End If
    End If
End Sub

Private Function CalculaMonto(nTotal As Double, nOtros As Double, nInafecto As Double) As Boolean
    CalculaMonto = True
    Dim entro As Boolean
    Dim nTC   As Double
    entro = False
    
        Dim ValorSaldo As Double
        Dim ValorMonto As Double
        
        'Convierte los montos soles a dolares
        If tdbMoneda.BoundText = gsMonedaNac And gintBiMoneda = 1 Then
            nTC = NE(tdbTC.Value)
            'If nTC = 0 Then nTC = 1
            ValorSaldo = Round(NE(tdbtSaldo.Text), 2)
            ValorMonto = Round((nTotal + nOtros + nInafecto) * nTC, 2)
            
        ElseIf tdbMoneda.BoundText = gsMonedaNac And gintBiMoneda = 0 Then
            ValorSaldo = Round(NE(tdbtSaldo.Text), 2)
            ValorMonto = Round(nTotal + nOtros + nInafecto, 2)
        Else
            nTC = NE(tdbTC.Value)
            'If nTC = 0 Then nTC = 1
            ValorSaldo = Round(NE(tdbtSaldo.Text), 2)
            ValorMonto = Round((nTotal + nOtros + nInafecto) * nTC, 2)
        End If
        
        If Round(ValorMonto, 2) > Round(ValorSaldo, 2) Then
            If tdbMoneda.BoundText = gsMonedaNac Then
                Mensajes "No se puede ingresar un monto mayor al saldo," & Salto(2) & "MONTO :  S/. " & Round(nTotal + nOtros + nInafecto, 2) & Salto(2) & "SALDO :  S/. " & lblSaldo, vbOKOnly + vbInformation
            Else
                Mensajes "No se puede ingresar un monto mayor al saldo," & Salto(2) & "MONTO :  US$ " & Round(nTotal + nOtros + nInafecto, 2) & " X " & Round(tdbTC.Value, 2) & " = S/." & Round(ValorMonto, 2) & Salto(2) & "SALDO :  S/. " & lblSaldo, vbOKOnly + vbInformation
            End If
            CalculaMonto = False
            
            'TdbTotal.Value = 0
            'Hallar_montos
            
        Else
            Hallar_montos
            ActivaUIT
        End If
    

End Function

Private Sub ManNuevo()
    If NE(nFilasVoucher) <= 0 Then
        Mensajes "Seleccione un registro del " & lblCabecera.Caption & Salto(1) & "para poder adicionar sus comprobantes de detalle", vbInformation
        Exit Sub
    End If


    Dim ValSaldo As Double
     ValSaldo = lblSaldo.Caption
    
    If op_compras.Value = True And NE(ValSaldo) <= 0 Then
        Mensajes "Ya no se pueden ingresar mas documentos, se cubrio el monto total del voucher", vbInformation
        Exit Sub
    End If
    
    If tdbcPeriodo.BoundText = "" Then
        Mensajes "Debe seleccionar un Periodo", vbInformation
        Exit Sub
    End If
    
    
    
    Dim Mes As String
    If gsPeriodo <> "" Then
        
        If gsPeriodo = "00" Then
            Mes = "01"
        End If
        If gsPeriodo > "12" Then
            Mes = "12"
        End If
        If gsPeriodo > "00" And gsPeriodo < "13" Then
            Mes = gsPeriodo
        End If
        
    End If
    

    
    dtpFechaBus = "01/" & Mes & "/" & gsAnio
    
    activar_botones False, False, False, True, True
    Me.SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    flag_grabar = True
    tdbIGV = nIGV
    Call Limpiar
    
    tdbMoneda.BoundText = tdbgAsientos.Columns(8)
    
    tdbtSaldo.Text = lblSaldo.Caption
    
    cboTipoDocReg.BoundText = CE(tdbgAsientos.Columns(12).Value)
    tdbSerie.Text = CE(tdbgAsientos.Columns(4).Value) 'BuscaUltimaSerie
    
    TdbNumDocReg.Text = BuscaUltimoDocumento
    
    
    tdbMoneda.Enabled = True
    
    Habilitar_frame_nombres True, True
    
    
    dtpFechaBus = tdbtFecha
    
    pSetFocus tdbSerie
    
End Sub
Private Function BuscaUltimaSerie() As String
    Dim sqlver  As String, valorDato As String
    Dim RegFiltro As String
    Dim RegFlag As String
    If op_compras.Value = True Then
        RegFiltro = lsLibroCom
        RegFlag = "C"
    Else
        RegFiltro = lsLibroVen
        RegFlag = "V"
    End If
    
    sqlver = "select max(Asd_cSeriedoc) from CNT_REG_BOLETAS a " & _
             "WHERE A.Emp_cCodigo = '" & gsEmpresa & "' AND A.PAN_CANIO = '" & gsAnio & "' AND A.Blta_cUserCrea <> '*' AND " & _
             "A.Per_cPeriodo = '" & tdbcPeriodo.BoundText & "' AND A.Tdo_cCodigo = '" & CE(tdbgAsientos.Columns(12).Value) & "' and blta_cflaglibro='" & RegFlag & "'"

    valorDato = ExtraeDescripcion(sqlver)
    If valorDato <> "" Then
        BuscaUltimaSerie = valorDato
    Else
        BuscaUltimaSerie = "001"
    End If


End Function

Private Function BuscaUltimoDocumento() As String
    Dim sqlver  As String, valorDato As String
    Dim RegFiltro As String
    Dim RegFlag As String
    If op_compras.Value = True Then
        RegFiltro = lsLibroCom
        RegFlag = "C"
    Else
        RegFiltro = lsLibroVen
        RegFlag = "V"
    End If
    
    sqlver = "select max(Asd_cNumdoc) from CNT_REG_BOLETAS a " & _
             "WHERE A.Emp_cCodigo = '" & gsEmpresa & "' AND A.PAN_CANIO = '" & gsAnio & "' AND A.Blta_cUserCrea <> '*' AND " & _
             "A.Per_cPeriodo = '" & tdbcPeriodo.BoundText & "' AND A.Tdo_cCodigo = '" & CE(tdbgAsientos.Columns(12).Value) & "' and blta_cflaglibro='" & RegFlag & "' and Asd_cSerieDoc = '" & CE(tdbSerie.Text) & "' "

    valorDato = ExtraeDescripcion(sqlver)
    If valorDato <> "" Then
        BuscaUltimoDocumento = Right("00000000" & NE(valorDato) + 1, 8)
    Else
        BuscaUltimoDocumento = "00000001"
    End If

End Function

Private Sub ManEliminar()
    If nFilas <= 0 Then
        Mensajes "Para eliminar primero seleccione un registro en el detalle", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    If MsgBox("Esta UD. seguro de Eliminar el Registro :  " & tdbgBoletas.Columns(5) & "", vbYesNo + vbInformation, "PLANILLAS") = vbYes Then
        Call Eliminar
        Call cargar_rs_boletas
        Call MuestraSaldo
    End If

End Sub

Private Sub ActivaUIT()
    Dim UIT As Double
    UIT = Val(NE(BuscaValorEnConfigOP("027"))) / 2
    If NE(TdbTotal.Value) >= NE(UIT) Then
        pSetFocus TdbTipoDoc
        On Error Resume Next
    End If
End Sub

Private Sub TdbTotal_LostFocus()
    Call Hallar_montos
    Call ActivaUIT
End Sub

Private Sub tlbOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim periodo As String
    Select Case Button.Index
        Case Is = "1" 'NUEVO
            Call ManNuevo
       Case Is = "2" 'VER
       Case Is = "3" 'GRABA
            Call GrabarDoc
       Case Is = "4" 'ELIMINA
            ManEliminar
       Case Is = "5" 'EDITAR
            ManModificar
            'pSetFocus tdbPeriodo
            
       Case Is = "6" 'IMPRIMIR
            FrmRepAuxiliarVentas.Show
            FrmRepAuxiliarVentas.ZOrder 0
            
       Case Is = "7" 'CANCELAR
            If SSTab1.Tab = 1 Then
                If MsgBox("Desea cancelar el registro del documento", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    Call CancelarDoc
                End If
                
                
            Else
                If MsgBox("Desea salir del mantenimiento de Registros Auxiliares", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    Unload Me
                End If
            End If
    End Select
End Sub

Private Sub ManModificar()
    
    If nFilas < 1 Then
        Mensajes "Seleccione un registro del detalle ", vbOKOnly + vbInformation
        Exit Sub
    End If
    dtpFechaBus = tdbtFecha
    
    activar_botones False, False, False, True, True
    
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    flag_grabar = False
    Call cargar_datos
    tlbOpciones.Buttons(5).Enabled = False
    tdbMoneda.Enabled = True
    
    Me.SSTab1.Tab = 1
    CargaInicial = False
    DoEvents
    ActivaUIT
    MuestraSaldo
    Hallar_montos
    
    pSetFocus tdbSerie
    
    CargaInicial = False
End Sub

Private Sub activar_botones(nuevo As Boolean, modificar As Boolean, Eliminar As Boolean, Cancelar As Boolean, Optional Salir As Boolean = True)
    tlbOpciones.Buttons(1).Enabled = nuevo
    tlbOpciones.Buttons(5).Enabled = modificar
    tlbOpciones.Buttons(4).Enabled = Eliminar
    
    tlbOpciones.Buttons(7).Enabled = Salir
    
End Sub
Private Function new_cod() As String
   Dim i As Integer
   Dim cad_num As String
   cad_num = ""
   If rs_boletas.RecordCount = 0 Then
      new_cod = "00000001"
   Else
         rs_boletas.MoveLast
         new_cod = Trim(Str(Val(rs_boletas(5)) + 1))
        
         If Len(new_cod) < 8 Then
          For i = Len(new_cod) To 7
          cad_num = cad_num & "0"
          Next
         End If
        new_cod = Trim(cad_num) & Trim(new_cod)
   End If
End Function
Private Sub cargar_datos()
    txtRegistro.Text = CE(tdbgBoletas.Columns(5))
    
    tdbSerie.Text = CE(tdbgBoletas.Columns(6))
    TdbNumDocReg = CE(tdbgBoletas.Columns(7))
    
    tdbMoneda.BoundText = CE(tdbgBoletas.Columns(25))
    
    Call tdbMoneda_ItemChange
    tdbTC.Value = IIf(CE(tdbgBoletas.Columns(24).Value) = "", 0, CE(tdbgBoletas.Columns(24).Value))
    
    If tdbMoneda.BoundText = gsMonedaNac Then
       TdbBaseImponible.Value = NE(tdbgBoletas.Columns(9).Value)
       TdbTotal.Value = NE(tdbgBoletas.Columns(12).Value)
       
    Else
       TdbBaseImponible.Value = NE(tdbgBoletas.Columns(15).Value)
       TdbTotal.Value = NE(tdbgBoletas.Columns(18).Value)
    End If
    
    cboTipoDocReg.BoundText = CE(tdbgBoletas.Columns(4).Value)
    TdbTipoDoc.BoundText = CE(tdbgBoletas.Columns(19).Value)
        
    TdbNroDoc.Text = CE(tdbgBoletas.Columns(20).Value)
    TdbNombres.Text = CE(tdbgBoletas.Columns(21).Value)
    TdbApellidos.Text = CE(tdbgBoletas.Columns(22).Value)
    
    If (CE(tdbgBoletas.Columns(31).Value) <> "ANULADO" And CE(tdbgBoletas.Columns(31).Value) <> vbNullString) Then
        tdbtEntidad.Text = Mid(CE(tdbgBoletas.Columns(31).Value), 1, InStr(1, CE(tdbgBoletas.Columns(31).Value), "-") - 1) 'InStr(1, CE(tdbgBoletas.Columns(31).Value), "-")
    End If
    
        
    If tdbMoneda.BoundText = gsMonedaNac Then
        tdbOtros = NE(tdbgBoletas.Columns(11).Value)
        tdbBaseImpInafecto = NE(tdbgBoletas.Columns(26).Value)
    Else
        tdbOtros = NE(tdbgBoletas.Columns(17).Value)
        tdbBaseImpInafecto = NE(tdbgBoletas.Columns(28).Value)
    End If
    
    chkInafecto = 0
    
    lblTotalDoc.Caption = Format(NE(TdbTotal) + NE(tdbOtros) + NE(tdbBaseImpInafecto), "###,###,##0.00")
End Sub

Private Sub Limpiar()

    tdbSerie.Text = ""
    TdbNumDocReg = ""
    tdbMoneda.BoundText = ""
    tdbTC.Value = 0
    TdbBaseImponible.Value = 0
    tdbBaseImpInafecto.Value = 0
    tdbOtros.Value = 0
    chkInafecto.Value = vbUnchecked
    TdbTotal.Value = 0
    TdbTipoDoc.BoundText = ""
    TdbNroDoc = ""
    TdbNombres = ""
    TdbApellidos = ""
    lblTotalDoc = "0.00"
    lblIGV = "0.00"
    
    lblConvBaseImp.Caption = "0.00"
    lblConvBaseImpInaf.Caption = "0.00"
    lblConvIgv.Caption = "0.00"
    lblConvOtros.Caption = "0.00"
    lblConvTotalDoc.Caption = "0.00"
    lblConvTotal.Caption = "0.00"
End Sub

'Private Sub txtTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        Tipo_Cambio = NE(tdbTC.Value)
'    End If
'End Sub

''--------------------------------------------------------------------------------
'' Project    :       Contabilidad
'' Procedure  :       lblnombre_LostFocus
'' Description:       Evento que se ejecuta al perder el enfoque el codigo de entidad
''
'' Parameters :
''--------------------------------------------------------------------------------
'Private Sub lblnombre_LostFocus()
'    If CE(lblnombre.Text) = "" Then
'        tdbtNombreEntidad.Text = ""
'    End If
'
'    fValidEntidad
'End Sub

''--------------------------------------------------------------------------------
'' Project    :       Contabilidad
'' Procedure  :       lblnombre_KeyDown
'' Description:       Evento que se ejecuta al presionar una tecla en el tipo de entidad
''
'' Parameters :       KeyCode (Integer)
''                    Shift (Integer)
''--------------------------------------------------------------------------------
'Private Sub lblnombre_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 112 Then
'        If tdbcTipoEntidad.BoundText = "" Then
'            Mensajes "Seleccione una entidad", vbOKOnly + vbInformation
'            pSetFocus tdbcTipoEntidad
'            Exit Sub
'        End If
'
'        Call LlamaBuscar(frmBuscador, Me.lblnombre.Name, lControl, "Entidad", Me, gsPeriodo, Me.tdbcTipoEntidad.BoundText)
'    End If
'
'
'End Sub

''--------------------------------------------------------------------------------
'' Project    :       Contabilidad
'' Procedure  :       lblnombre_Change
'' Description:       Evento que se ejecuta al cambiar el tipo de entidad
''
'' Parameters :
''--------------------------------------------------------------------------------
'Private Sub lblnombre_Change()
'    If CE(lblnombre.Text) = "" Then
'        tdbtNombreEntidad.Text = ""
'    End If
'End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtEntidad_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el tipo de entidad
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cCodigo As String
    cCodigo = ""
    
    If KeyCode = 112 Then
        cCodigo = BuscarEntidad(Me, "C", "")
    
        If cCodigo <> "" Then
            tdbtEntidad.Text = cCodigo
            TdbNombres.Text = BuscaNombreEntidad("C", cCodigo)
            TdbTipoDoc.BoundText = gsCampo3
            TdbNroDoc.Text = gsCampo4
            TdbApellidos.Text = gsCampo5
        End If
        
    End If
    
    
End Sub

