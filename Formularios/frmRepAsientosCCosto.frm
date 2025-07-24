VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepAsientosCCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Analisis por Centro de Costo"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmRepAsientosCCosto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   7695
   Begin VB.Frame fraTodo 
      Height          =   3165
      Left            =   45
      TabIndex        =   6
      Top             =   -45
      Width           =   7530
      Begin VB.Frame Frame1 
         Height          =   2160
         Left            =   225
         TabIndex        =   7
         Top             =   180
         Width           =   7050
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   1920
            Left            =   135
            TabIndex        =   8
            Top             =   180
            Width           =   6795
            Begin TDBDate6Ctl.TDBDate dtpDesde 
               Height          =   300
               Left            =   1080
               TabIndex        =   1
               Tag             =   "enabled"
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   529
               Calendar        =   "frmRepAsientosCCosto.frx":0ECA
               Caption         =   "frmRepAsientosCCosto.frx":0FCC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAsientosCCosto.frx":1030
               Keys            =   "frmRepAsientosCCosto.frx":104E
               Spin            =   "frmRepAsientosCCosto.frx":10BA
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
            Begin TDBDate6Ctl.TDBDate dtpHAsta 
               Height          =   300
               Left            =   1080
               TabIndex        =   2
               Tag             =   "enabled"
               Top             =   1080
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   529
               Calendar        =   "frmRepAsientosCCosto.frx":10E2
               Caption         =   "frmRepAsientosCCosto.frx":11E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAsientosCCosto.frx":1248
               Keys            =   "frmRepAsientosCCosto.frx":1266
               Spin            =   "frmRepAsientosCCosto.frx":12D2
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
            Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
               Height          =   315
               Left            =   1080
               TabIndex        =   0
               Tag             =   "_"
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   556
               Caption         =   "frmRepAsientosCCosto.frx":12FA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAsientosCCosto.frx":1366
               Key             =   "frmRepAsientosCCosto.frx":1384
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
               Format          =   "aA"
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   12
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
            Begin TDBText6Ctl.TDBText tdbtDescripcionDesde 
               Height          =   315
               Left            =   2640
               TabIndex        =   9
               Tag             =   "_"
               Top             =   360
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
               _ExtentY        =   556
               Caption         =   "frmRepAsientosCCosto.frx":13C8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAsientosCCosto.frx":1434
               Key             =   "frmRepAsientosCCosto.frx":1452
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   1
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
            Begin TrueOleDBList70.TDBCombo tdbcMoneda 
               Height          =   300
               Left            =   1080
               TabIndex        =   3
               Tag             =   "_"
               Top             =   1440
               Width           =   1575
               _ExtentX        =   2778
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
               _PropDict       =   $"frmRepAsientosCCosto.frx":14A4
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
            Begin VB.Label lblHasta 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hasta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   1125
               Width           =   495
            End
            Begin VB.Label lblDesde 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desde"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   240
               TabIndex        =   13
               Top             =   765
               Width           =   555
            End
            Begin VB.Label lblCCosto 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CCosto"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   240
               TabIndex        =   12
               Top             =   405
               Width           =   630
            End
            Begin VB.Label lblMoneda 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
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
               Left            =   240
               TabIndex        =   11
               Top             =   1485
               Width           =   660
            End
            Begin VB.Label lblSiNo 
               Alignment       =   2  'Center
               Caption         =   "Si no se indica el centro de Costo, entonces mostrará todos los Centros de Costo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   3240
               TabIndex        =   10
               Top             =   840
               Width           =   2775
            End
         End
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   3780
         TabIndex        =   5
         Top             =   2565
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepAsientosCCosto.frx":152B
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   2070
         TabIndex        =   4
         Top             =   2565
         Width           =   1665
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepAsientosCCosto.frx":1AC5
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
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
      TabIndex        =   15
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepAsientosCCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdImprimir_Click()
    ' *** Abrir el reporte y enviar los parametros
    Dim Tipo As String
    Dim matriz_fecha(8) As Variant
    
    Screen.MousePointer = vbHourglass
    
    Tipo = ""
    matriz_fecha(0) = "@Tipo;" & Tipo & ";True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz_fecha(3) = "@desde;" & dtpDesde.Text & ";True"
    matriz_fecha(4) = "@hasta;" & dtpHasta.Text & ";True"
    matriz_fecha(5) = "@Cos_cCodigo;" & tdbtCuentaDesde & ";True"
    matriz_fecha(6) = "@moneda;" & Me.tdbcMoneda.BoundText & ";True"
    
    matriz_fecha(7) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(8) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(4) As Variant
    
'    If Me.optGrupo01.Value = True Then
'        formulas(0) = "grupo01 = {spCn_RptAnalisisDocumentos;1.TITULO}"
'        formulas(1) = "grupo02 = {spCn_RptAnalisisDocumentos;1.Cos_cCodigo}"
'    End If
'    If Me.optGrupo02.Value = True Then
'        formulas(0) = "grupo01 = {spCn_RptAnalisisDocumentos;1.Cos_cCodigo}"
'        formulas(1) = "grupo02 = {spCn_RptAnalisisDocumentos;1.Pla_cCuentaContable}"
'    End If
'    If Me.optGrupo03.Value = True Then
'        formulas(0) = "grupo01 = {spCn_RptAnalisisDocumentos;1.Pla_cCuentaContable}"
'        formulas(1) = "grupo02 = {spCn_RptAnalisisDocumentos;1.Cos_cCodigo}"
'    End If
'
'    If Me.optOrden1.Value = True Then
'        formulas(2) = "orden01 = {spCn_RptAnalisisDocumentos;1.Ase_dFecha}"
'        formulas(3) = "orden02 = {spCn_RptAnalisisDocumentos;1.Lib_cTipoLibro}"
'    Else
'        formulas(2) = "orden01 = {spCn_RptAnalisisDocumentos;1.Lib_cTipoLibro}"
'        formulas(3) = "orden02 = {spCn_RptAnalisisDocumentos;1.Asd_dFecDoc}"
'    End If
'    If chkDestino.Value = "1" Then
'        formulas(4) = "conDestino = '*'"
'    Else
'        formulas(4) = "conDestino = '0'"
'    End If
    

    AbreReporteParam gsDSN, Me, rutaReportes & "RptAsientosxCentroCosto.rpt", crptToWindow, "Reporte de Asientos por Centro de Costo", "", matriz_fecha(), formulas()
    
    ' ***
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim VarMes As String
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    Call LlenaCombos
    Call BuscarMonedaNacional
     If gsPeriodo = "00" Then
        VarMes = "01/01/" + gsAnio
        dtpHasta = UltimoDiaMes("01", gsAnio)
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        VarMes = "01/12/" + gsAnio
        dtpHasta = UltimoDiaMes("12", gsAnio)
    Else
        VarMes = "01/" & gsPeriodo & "/" + gsAnio
        dtpHasta = UltimoDiaMes(gsPeriodo, gsAnio)
    End If
    dtpDesde = VarMes
    
    ActivarControl tdbtDescripcionDesde, False
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
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
    Set frmRepAsientosCCosto = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
            Case "CentroCosto" '
                tdbtCuentaDesde = Trim(param0)
                tdbtDescripcionDesde = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtCuentaDesde
    End Select
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
       Call LlamaBuscar(frmBuscador, "CentroCosto", Control, "CentroCostoN", Me, gsPeriodo)
    End If
End Sub

Private Function fValidCCosto() As Boolean
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As New ADODB.Recordset, arrDatos() As Variant
    Dim sqlver As String
    
    fValidCCosto = False
    
    tdbtDescripcionDesde = ""
    
    sqlver = "SELECT Cos_cDescripcion, Cos_cTitulo From CNT_CENTRO_COSTO " & _
             "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Cos_cDeleted <> '*' AND Cos_cEstado = 'A'" & _
             "AND Cos_cCodigo = '" & tdbtCuentaDesde & "' "
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State > 0 Then
       If rsArreglo("Cos_cTitulo") = "S" Then
          Mensajes "Centro de Costo esta definido como titulo, revisar", vbInformation
          Exit Function
       End If
       tdbtDescripcionDesde = rsArreglo("Cos_cDescripcion")
       fValidCCosto = True
    Else
       Mensajes "Codigo de Centro de Costo no existe, verificar...", vbInformation
    End If
End Function

Private Sub tdbtCuentaDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not fValidCCosto Then
          pSetFocus tdbtCuentaDesde
       Else
          pSendKeys "{tab}"
       End If
    End If
End Sub
