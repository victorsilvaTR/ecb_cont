VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Begin VB.Form frmManPlantillaTipoAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tipo de Asiento"
   ClientHeight    =   5955
   ClientLeft      =   2610
   ClientTop       =   2940
   ClientWidth     =   7965
   Icon            =   "frmManPlantillaTipoAsiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7965
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5310
      Left            =   45
      TabIndex        =   11
      Top             =   510
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   9366
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consultas de Tipo de Asientos"
      TabPicture(0)   =   "frmManPlantillaTipoAsiento.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Tipo de Asientos"
      TabPicture(1)   =   "frmManPlantillaTipoAsiento.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4875
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   7665
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   3510
            Left            =   180
            TabIndex        =   2
            Top             =   1215
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   6191
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Libro"
            Columns(0).DataField=   "Lib_cDescripcion"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Diario"
            Columns(1).DataField=   "Lib_cTipoLibro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Codigo "
            Columns(2).DataField=   "Asl_cOperacion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Descripción "
            Columns(3).DataField=   "Asl_cDescripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3916"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3836"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Merge=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=767"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=688"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8724"
            Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(17)=   "Column(2).Width=1640"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1561"
            Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(23)=   "Column(3).Width=5318"
            Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=5239"
            Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Named:id=33:Normal"
            _StyleDefs(55)  =   ":id=33,.parent=0"
            _StyleDefs(56)  =   "Named:id=34:Heading"
            _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   ":id=34,.wraptext=-1"
            _StyleDefs(59)  =   "Named:id=35:Footing"
            _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(61)  =   "Named:id=36:Selected"
            _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(63)  =   "Named:id=37:Caption"
            _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(65)  =   "Named:id=38:HighlightRow"
            _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1770
            TabIndex        =   0
            Top             =   510
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   556
            Caption         =   "frmManPlantillaTipoAsiento.frx":0F02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlantillaTipoAsiento.frx":0F6E
            Key             =   "frmManPlantillaTipoAsiento.frx":0F8C
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
         Begin TDBText6Ctl.TDBText tdbtLibroBus 
            Height          =   315
            Left            =   1770
            TabIndex        =   1
            Top             =   840
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "frmManPlantillaTipoAsiento.frx":0FDE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlantillaTipoAsiento.frx":104A
            Key             =   "frmManPlantillaTipoAsiento.frx":1068
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
            Index           =   0
            Left            =   480
            TabIndex        =   19
            Top             =   885
            Width           =   420
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
            Index           =   5
            Left            =   480
            TabIndex        =   18
            Top             =   555
            Width           =   990
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
            Left            =   480
            TabIndex        =   17
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   -74730
         TabIndex        =   12
         Top             =   540
         Width           =   6930
         Begin VB.Frame Frame3 
            Caption         =   "Detalle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Left            =   165
            TabIndex        =   22
            Top             =   1230
            Width           =   6540
            Begin VB.CommandButton cmdEliminarDestino 
               Caption         =   " Eliminar Item"
               Enabled         =   0   'False
               Height          =   510
               Left            =   4995
               Picture         =   "frmManPlantillaTipoAsiento.frx":10BA
               Style           =   1  'Graphical
               TabIndex        =   10
               Tag             =   "enabled"
               Top             =   1350
               Width           =   1350
            End
            Begin VB.CommandButton cmdInsertar 
               Caption         =   "Insertar Item"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   5025
               Picture         =   "frmManPlantillaTipoAsiento.frx":1644
               Style           =   1  'Graphical
               TabIndex        =   9
               Tag             =   "_"
               Top             =   705
               Width           =   1350
            End
            Begin VB.ComboBox cmbTipo 
               Height          =   315
               ItemData        =   "frmManPlantillaTipoAsiento.frx":1BCE
               Left            =   2850
               List            =   "frmManPlantillaTipoAsiento.frx":1BD8
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   270
               Width           =   1230
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnPorc 
               Height          =   315
               Left            =   5325
               TabIndex        =   8
               Top             =   240
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   556
               Calculator      =   "frmManPlantillaTipoAsiento.frx":1BE9
               Caption         =   "frmManPlantillaTipoAsiento.frx":1C09
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlantillaTipoAsiento.frx":1C75
               Keys            =   "frmManPlantillaTipoAsiento.frx":1C93
               Spin            =   "frmManPlantillaTipoAsiento.frx":1CEB
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00 %"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00 %"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   100
               MaxValueVT      =   1802698757
               MinValueVT      =   1769209861
            End
            Begin TrueOleDBList70.TDBList tdblDestino 
               Height          =   2340
               Left            =   225
               TabIndex        =   23
               Top             =   705
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   4128
               _LayoutType     =   4
               _RowHeight      =   -2147483647
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
               Columns(2).Caption=   "Nombre Cuenta"
               Columns(2).DataField=   "Nombre Cuenta"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "%"
               Columns(3).DataField=   "%"
               Columns(3).NumberFormat=   "Standard"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   4
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=4"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=1826"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1746"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(13)=   "Column(2).Width=3995"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3916"
               Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(19)=   "Column(3).Width=688"
               Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=609"
               Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
               Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&"
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
               _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
               _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
               _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
               _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
               _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
               _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
               _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
               _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
               _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
               _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
               _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(49)  =   "Named:id=33:Normal"
               _StyleDefs(50)  =   ":id=33,.parent=0"
               _StyleDefs(51)  =   "Named:id=34:Heading"
               _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(53)  =   ":id=34,.wraptext=-1"
               _StyleDefs(54)  =   "Named:id=35:Footing"
               _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(56)  =   "Named:id=36:Selected"
               _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(58)  =   "Named:id=37:Caption"
               _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(60)  =   "Named:id=38:HighlightRow"
               _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(62)  =   "Named:id=39:EvenRow"
               _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(64)  =   "Named:id=40:OddRow"
               _StyleDefs(65)  =   ":id=40,.parent=33"
               _StyleDefs(66)  =   "Named:id=41:RecordSelector"
               _StyleDefs(67)  =   ":id=41,.parent=34"
               _StyleDefs(68)  =   "Named:id=42:FilterBar"
               _StyleDefs(69)  =   ":id=42,.parent=33"
            End
            Begin TDBText6Ctl.TDBText tdbtCtaDestino 
               Height          =   315
               Left            =   960
               TabIndex        =   6
               Tag             =   "_"
               Top             =   255
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   556
               Caption         =   "frmManPlantillaTipoAsiento.frx":1D13
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlantillaTipoAsiento.frx":1D7F
               Key             =   "frmManPlantillaTipoAsiento.frx":1D9D
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
               Format          =   "9"
               FormatMode      =   0
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Porcentaje"
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
               Left            =   4305
               TabIndex        =   26
               Top             =   285
               Width           =   885
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Tipo"
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
               Left            =   2355
               TabIndex        =   25
               Top             =   300
               Width           =   360
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta"
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
               Left            =   210
               TabIndex        =   24
               Top             =   285
               Width           =   600
            End
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   5655
            TabIndex        =   4
            Top             =   480
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   556
            Caption         =   "frmManPlantillaTipoAsiento.frx":1DEF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlantillaTipoAsiento.frx":1E5B
            Key             =   "frmManPlantillaTipoAsiento.frx":1E79
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
            Left            =   1365
            TabIndex        =   5
            Tag             =   "_"
            Top             =   885
            Width           =   5130
            _Version        =   65536
            _ExtentX        =   9049
            _ExtentY        =   556
            Caption         =   "frmManPlantillaTipoAsiento.frx":1ECB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlantillaTipoAsiento.frx":1F37
            Key             =   "frmManPlantillaTipoAsiento.frx":1F55
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
            MaxLength       =   60
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
         Begin TrueOleDBList70.TDBCombo tdbcLibro 
            Height          =   300
            Left            =   1365
            TabIndex        =   3
            Tag             =   "enabled"
            Top             =   510
            Width           =   3120
            _ExtentX        =   5503
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
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
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
            Splits(0)._ColumnProps(17)=   "Column(3).Width=2196"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2117"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2196"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2117"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
            _PropDict       =   $"frmManPlantillaTipoAsiento.frx":1FA7
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
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtNombreDestino 
            Height          =   315
            Left            =   480
            TabIndex        =   21
            Top             =   1980
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   556
            Caption         =   "frmManPlantillaTipoAsiento.frx":202E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlantillaTipoAsiento.frx":209A
            Key             =   "frmManPlantillaTipoAsiento.frx":20B8
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Diario"
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
            Left            =   300
            TabIndex        =   20
            Top             =   555
            Width           =   495
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
            Index           =   3
            Left            =   300
            TabIndex        =   15
            Top             =   915
            Width           =   990
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
            Index           =   8
            Left            =   4590
            TabIndex        =   14
            Top             =   555
            Width           =   600
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
            Left            =   270
            TabIndex        =   13
            Top             =   210
            Width           =   1515
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   3615
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
            Picture         =   "frmManPlantillaTipoAsiento.frx":210A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":24E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":28BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":2C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":3072
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":344C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":3826
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":3C00
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
            Picture         =   "frmManPlantillaTipoAsiento.frx":4C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":4D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":4ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":5028
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":5182
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":52DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":5436
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":5590
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":56EA
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
            Picture         =   "frmManPlantillaTipoAsiento.frx":5844
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":5DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":6378
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":6912
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":6EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":7446
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":79E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":7F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlantillaTipoAsiento.frx":8514
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   27
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
Attribute VB_Name = "frmManPlantillaTipoAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
'Dim lrstabla As Recordset
Dim lrsTabla As ADODB.Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim Control As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmbTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub

Private Sub cmdEliminarDestino_Click()
    Dim Fila As Integer
    Dim i As Integer
    
    If tdblDestino.ListCount > 0 Then
        Fila = tdblDestino.Row
        If tdblDestino.Row <> -1 Then
            tdblDestino.RemoveItem Fila
            tdblDestino.ReBind
        Else
            Mensajes "Seleccione el item a eliminar", vbInformation
        End If
    End If
End Sub

Private Sub cmdInsertar_Click()
    ' *** Q el codigo sea diferente de nada
    If CE(tdbtCtaDestino.Text) = "" Then
        Mensajes "Ingrese una cuenta", vbInformation
        Exit Sub
    End If
    If tdbnPorc = 0 Then
        Mensajes "Ingrese una cantidad en porcentaje diferente a 0", vbInformation
        Exit Sub
    End If
    ' *** Insertar Registro a la Lista de Destino
    tdblDestino.AddItem Mid$(Trim(Me.cmbTipo), 1, 1) & "; " & tdbtCtaDestino & "; " & tdbtNombreDestino & " ;" & tdbnPorc
    
    ' *** Limpiar lo q se escribio previamente
    pSetFocus tdbtCtaDestino
    Me.tdbtCtaDestino = ""
    Me.tdbtNombreDestino = ""
    tdbnPorc = 100
    ' ***
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left + 15 - 300
            .Height = Me.Height - .Top + 15 - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 500
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        '*** REDIMENSIONAR DETALLE
        Frame2.Height = Frame1.Height
        Frame2.Width = Frame1.Width
        
        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
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
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
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
        ' ***
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            ReDim lArrMnt(11) As Variant
            lArrMnt(0) = "ELIMINAR"     ' Accion
            lArrMnt(1) = gsEmpresa      ' Codigo de Empresa
            lArrMnt(2) = gsAnio         ' Codigo de Plantilla
            lArrMnt(3) = tdbgCostos.Columns(1).Value    ' Tipo de Libro
            lArrMnt(4) = tdbgCostos.Columns(2).Value    ' Codigo Operacion
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoAsiento", lArrMnt(), True) = False Then
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
    Call CargaDatosRegistro
    If lRegElim = False Then
        lblMante = "VER REGISTRO"
        SSTCentroCosto.TabEnabled(1) = True
        SSTCentroCosto.TabEnabled(0) = False
        SSTCentroCosto.Tab = 1
        tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(7).Image = 8
        lTipoMnt = "EDITAR"
        Call AseguraControl(Me, True)
    Else
        lRegElim = False
    End If
    
    tdbcLibro.Enabled = False
End Sub

Private Sub Editar()
    Call CargaDatosRegistro
    
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
        pSetFocus tdbtDescripcion
    Else
        lRegElim = False
    End If
    
    tdbcLibro.Enabled = False
End Sub

Private Sub ManNuevo()
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    
    tdbcLibro.Enabled = True
    
    pSetFocus tdbcLibro
    tdbtCodigo.ReadOnly = True
    tdblDestino.Clear
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
    If Me.lblMante = "VER REGISTRO" Then
        Call AseguraControl(Me, False)
    Else
        Call HabilitaControl(Me)
    End If
    Call TabMantenimiento(False)
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
End Sub

Private Sub Imprimir()
    Dim matriz(15) As Variant
    Dim Titulo As String
    Titulo = "Tipos de Asiento"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;LIBRO - DESCRIPCION;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;;True"
    matriz(5) = "@Titulo04;;True"
    matriz(6) = "@Titulo05;DEBE/HABER;True"
    matriz(7) = "@Titulo06;CUENTA;True"
    matriz(8) = "@Titulo07;PORCENT.;True"
    matriz(9) = "@Tipo;TIPO_ASIENTO;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"
    
    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandarAgrupado.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub

Private Function ValidaDestinoPorc() As Boolean
    Dim i As Integer
    Dim Suma As Double
    Dim SumaHaber As Double
    ValidaDestinoPorc = False
    Suma = 0
    SumaHaber = 0
    
    For i = 0 To tdblDestino.ListCount - 1
        tdblDestino.Row = i
        If tdblDestino.Columns(0) = "D" Then
            Suma = Suma + tdblDestino.Columns(3).Value
        End If
        
        If tdblDestino.Columns(0) = "H" Then
            SumaHaber = SumaHaber + tdblDestino.Columns(3).Value
        End If
        
    Next i
    
    If CDec(Suma) <> CDec(SumaHaber) Then
        Mensajes "Los porcentajes acumulados del DEBE y del HABER deben ser los mismos", vbOKOnly + vbInformation
        Exit Function
    End If
    
    
    If Suma = 0 And SumaHaber = 0 Then
        Mensajes "Ingrese las cuentas del asiento automatico", vbOKOnly + vbInformation
        Exit Function

    End If
    
    ValidaDestinoPorc = True
End Function

Private Sub Grabar()
    If validarDatos = False Then Exit Sub
    If ValidaDestinoPorc = False Then Exit Sub
    


    Dim clsMante As clsMantoTablas
    Dim i As Integer
    If lTipoMnt = "INSERTAR" Then Call GeneraCodigo
    
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    If tdblDestino.ListCount > 0 Then
        For i = 0 To tdblDestino.ListCount - 1
            Call CargaArregloMnt(i)
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoAsiento", lArrMnt(), False) = False Then
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                Exit Sub
            End If
        Next
        lArrMnt(0) = "ELIMINAMAYOR"           ' Accion
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoAsiento", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Exit Sub
        End If
    End If
    Call Cancelar
    CargaTabla
    ' ***
    Mensajes "Los datos se grabaron con exito...", vbInformation
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If CE(tdbtCodigo.Text) = "" Then
        Mensajes "Ingrese el codigo", vbInformation
        pSetFocus tdbtCodigo
        Exit Function
    End If
    
    If CE(tdbtDescripcion.Text) = "" Then
        Mensajes "Ingrese la descripcion", vbInformation
        pSetFocus tdbtDescripcion
        Exit Function
    End If
    
    
    If tdblDestino.ListCount < 0 Then
        Mensajes "Ingrese detalle al tipo de Asiento", vbInformation
        pSetFocus Me.tdbtCtaDestino
        Exit Function
    End If
    ' ***
    validarDatos = True
End Function

Private Sub CargaArregloMnt(Numero As Integer)
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(11) As Variant
    lArrMnt(0) = lTipoMnt           ' Accion
    lArrMnt(1) = gsEmpresa          ' Empresa
    lArrMnt(2) = gsAnio         ' Año
    lArrMnt(3) = Me.tdbcLibro.BoundText    ' Libro
    lArrMnt(4) = Me.tdbtCodigo                ' Codigo
    lArrMnt(5) = Numero + 1         ' item
    lArrMnt(6) = Me.tdbtDescripcion          ' descripcion
    tdblDestino.Row = Numero
    lArrMnt(7) = Trim(tdblDestino.Columns(0).Value) ' Tipomov
    lArrMnt(8) = Trim(tdblDestino.Columns(1).Value) ' Cuenta
    lArrMnt(9) = tdblDestino.Columns(3).Value       ' Porcentaje
    lArrMnt(10) = "A"
    lArrMnt(11) = gsUsuario
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub LlenaCombos()
    Dim sqlcombos As String
    
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND " & _
                "(LIB_CTIPOLIBRO='" & lsLibroCom & "' OR LIB_CTIPOLIBRO='" & lsLibroVen & "' ) " & _
                "ORDER BY LIB_CDESCRIPCION "
                '"LIB_CTIPOLIBRO='" & lsLibroCajIng & "' OR LIB_CTIPOLIBRO='" & lsLibroCajEgr & "' OR " & _
                '"LIB_CTIPOLIBRO='" & lsLibroDiario & "') " & _

    LlenarComboAddItem tdbcLibro, sqlcombos

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)

    Call pCargaCfgLibro
    Call LlenaCombos
    Call CargaTabla
    
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    SSTCentroCosto.Tab = 0
    cmbTipo.ListIndex = 0
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    Call CerrarRecordSet(lrsTabla)
    
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    'Set lrstabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaTipoAsiento 'SEL_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '', '', 0, '', '', '', 0, '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos)
    If Not lrsTabla Is Nothing Then
        ' *** Llenar grilla con el RecordSet
        lrsTabla.Sort = "Lib_cTipoLibro, Asl_cOperacion"
        tdbgCostos.DataSource = lrsTabla
        ' ***
        Exit Sub
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_GrabaTipoAsiento 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbgCostos.Columns(1).Value & "', '" & tdbgCostos.Columns(2).Value & "', 0, '', '', '', 0, '', '' "
    'sqlSp = "spCn_GrabaTipoEntidad 'SEL_REG', '" & gsEmpresa & "', '" & tdbgCostos.Columns(0).Value & "', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro el registro. Probablemente eliminado desde otra sesion", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    ' *** Asignando Datos al tipo
    
    tdbcLibro.BoundText = CE(rsArreglo!Lib_cTipoLibro)
    tdbtCodigo = CE(rsArreglo!Asl_cOperacion)
    tdbtDescripcion = CE(rsArreglo!Asl_cDescripcion)
    'Asl_nSecuencia, AslTipoMov, Asl_cCuenta, Asl_nPorcen,
    tdblDestino.Clear
    Do While Not rsArreglo.EOF
        tdblDestino.AddItem rsArreglo!AslTipoMov & "; " & rsArreglo!Asl_cCuenta & "; " & _
        rsArreglo!Pla_cNombreCuenta & " ;" & rsArreglo!Asl_nPorcen
        'rsArreglo!Pla_cNombreCuenta & " ;" & rsArreglo!Asl_nPorcen
        rsArreglo.MoveNext
    Loop
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(1) As String
    Dim i As Integer
    
    If lrsTabla Is Nothing Then Exit Sub
    On Local Error GoTo ErrorEjecucion
    cadena = ""
    If Trim(tdbtDescripcionBus) <> "" Then filtros(0) = "Asl_cDescripcion like '*" & tdbtDescripcionBus & "*'"
    If Trim(tdbtLibroBus) <> "" Then filtros(1) = "Lib_cTipoLibro like '*" & tdbtLibroBus & "*'"
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
    Exit Sub
ErrorEjecucion:
    If Err.Number <> "3265" Then Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub tdbcLibro_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbcLibro.Locked = True
    Else
        tdbcLibro.Locked = False
    End If
    ' ***
End Sub

Private Sub tdbcLibro_ItemChange()
    ' *** Autogenerar Codigo
    If lTipoMnt = "INSERTAR" Then Call GeneraCodigo
    ' ***
End Sub

Private Sub GeneraCodigo()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim lrsCodigo As New ADODB.Recordset
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaTipoAsiento 'SEL_COR', '" & gsEmpresa & "', '" & gsAnio & "', '" & Me.tdbcLibro.BoundText & "', '', 0, '', '', '', 0, '', '' "
    arrDatos = Array(sqlSp)
    Set lrsCodigo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsCodigo Is Nothing Then
        ' *** Hallar Codigo
        Me.tdbtCodigo = lrsCodigo(0).Value
        Exit Sub
    End If
    
    Set lrsCodigo = Nothing
End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
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

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtCodigo = Replace(tdbtCodigo, "'", "")
       tdbtCodigo.SelStart = Len(tdbtCodigo)
    End If
End Sub

Private Sub tdbtCtaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCtaDestino.Name, Control, "CuentasN", Me, gsPeriodo, tdbtCtaDestino)
End Sub

Private Sub tdbtCtaDestino_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCtaDestino <> "" And Me.Enabled = True Then
        tdbtNombreDestino = ExisteCtaNoTitulo(tdbtCtaDestino, "N")
        If SSTCentroCosto.TabEnabled(1) = True Then
           If tdbtNombreDestino.Text = "" Then
              tdbtCtaDestino.Text = ""
              pSetFocus tdbtCtaDestino
           End If
        End If
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
            Case "tdbtCtaDestino"   ' *** Caso de cliente
                tdbtCtaDestino = Trim(param0)
                Me.tdbtNombreDestino = Trim(param1)
                Unload frmBuscador
                pSetFocus cmbTipo
    End Select
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

Private Sub tdbtCodigo_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbtCodigo.ReadOnly = True
    Else
        tdbtCodigo.ReadOnly = False
    End If
    ' ***
End Sub

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

Private Sub tdbtLibroBus_Change()
    
    If gsKey = 219 Then
       tdbtLibroBus = Replace(tdbtLibroBus, "'", "")
       tdbtLibroBus.SelStart = Len(tdbtLibroBus)
    End If
    
    Call FiltrarRecordSet
End Sub

Private Sub tdbtLibroBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub
