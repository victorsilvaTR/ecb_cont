VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmManRegActRenFin 
   Caption         =   " ::: Registro de Activos en Arrendamiento Financiero :::"
   ClientHeight    =   5565
   ClientLeft      =   540
   ClientTop       =   855
   ClientWidth     =   13860
   Icon            =   "FrmManRegActRenFin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   13860
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   13875
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
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
         _PropDict       =   $"FrmManRegActRenFin.frx":0ECA
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
      Begin TDBDate6Ctl.TDBDate dtpFechaDoc 
         Height          =   225
         Left            =   4215
         TabIndex        =   2
         Tag             =   "enabled"
         Top             =   165
         Visible         =   0   'False
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   397
         Calendar        =   "FrmManRegActRenFin.frx":0F51
         Caption         =   "FrmManRegActRenFin.frx":1053
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegActRenFin.frx":10B7
         Keys            =   "FrmManRegActRenFin.frx":10D5
         Spin            =   "FrmManRegActRenFin.frx":1129
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010185729
         Value           =   38974
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber tdbtPorcx 
         Height          =   375
         Left            =   6690
         TabIndex        =   3
         Top             =   375
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "FrmManRegActRenFin.frx":1151
         Caption         =   "FrmManRegActRenFin.frx":1171
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegActRenFin.frx":11D5
         Keys            =   "FrmManRegActRenFin.frx":11F3
         Spin            =   "FrmManRegActRenFin.frx":122D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0.00"
         EditMode        =   2
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   1
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
         ShowContextMenu =   1
         ValueVT         =   1703937
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbtDepAcu 
         Height          =   375
         Left            =   6555
         TabIndex        =   4
         Top             =   -165
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "FrmManRegActRenFin.frx":1255
         Caption         =   "FrmManRegActRenFin.frx":1275
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegActRenFin.frx":12D9
         Keys            =   "FrmManRegActRenFin.frx":12F7
         Spin            =   "FrmManRegActRenFin.frx":1331
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0.00"
         EditMode        =   2
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   1
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
         ShowContextMenu =   1
         ValueVT         =   1703937
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbtVU 
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "FrmManRegActRenFin.frx":1359
         Caption         =   "FrmManRegActRenFin.frx":1379
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegActRenFin.frx":13DD
         Keys            =   "FrmManRegActRenFin.frx":13FB
         Spin            =   "FrmManRegActRenFin.frx":1435
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   2
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   1
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
         ShowContextMenu =   1
         ValueVT         =   1703937
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TrueOleDBGrid70.TDBGrid tdbgActivo 
         Height          =   4230
         Left            =   1455
         TabIndex        =   6
         Top             =   1065
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   7461
         _LayoutType     =   4
         _RowHeight      =   18
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Ase_cNummov"
         Columns(0).DataField=   "Ase_cNummov"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Ase_nVoucher"
         Columns(1).DataField=   "Ase_nVoucher"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha del Contrato"
         Columns(2).DataField=   "Act_dFechaContrato"
         Columns(2).NumberFormat=   "External Editor"
         Columns(2).ExternalEditor=   "dtpFechaDoc"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo de Bien"
         Columns(3).DataField=   "Pla_cNombreCuenta"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Moneda"
         Columns(4).DataField=   "Mon_cNombreLargo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Fecha de Adquisicion"
         Columns(5).DataField=   "Asd_dFecDoc"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Pla_cCuentaContable"
         Columns(6).DataField=   "Pla_cCuentaContable"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Vida Util"
         Columns(7).DataField=   "Act_nVidaUtil"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Valor del Bien"
         Columns(8).DataField=   "Asd_nDebeSoles"
         Columns(8).NumberFormat=   "External Editor"
         Columns(8).ExternalEditor=   "tdbtPorc"
         Columns(8).ExternalEditor.vt=   8
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Asd_nHaberSoles"
         Columns(9).DataField=   "Asd_nHaberSoles"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Ten_cTipoEntidad"
         Columns(10).DataField=   "Ten_cTipoEntidad"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Ent_cCodEntidad"
         Columns(11).DataField=   "Ent_cCodEntidad"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Asd_cTipoDoc"
         Columns(12).DataField=   "Asd_cTipoDoc"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Asd_cSerieDoc"
         Columns(13).DataField=   "Asd_cSerieDoc"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Asd_cNumDoc"
         Columns(14).DataField=   "Asd_cNumDoc"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Asd_cTipoMoneda"
         Columns(15).DataField=   "Asd_cTipoMoneda"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Depreciacion %"
         Columns(16).DataField=   "Act_nDepreciacionPorc"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Depreciacion Acumulada"
         Columns(17).DataField=   "Act_nDepreciacionAcumulada"
         Columns(17).NumberFormat=   "External Editor"
         Columns(17).ExternalEditor=   "tdbtDepAcu"
         Columns(17).ExternalEditor.vt=   8
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Depreciacion del Ejercicio"
         Columns(18).DataField=   "Act_nDepreciacionEjercicio"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Total Depreciacion"
         Columns(19).DataField=   ""
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "Valor Neto"
         Columns(20).DataField=   ""
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   21
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).ScrollGroup=   0
         Splits(0).Size  =   0
         Splits(0).Size.vt=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   16711680
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=21"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=532"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=532"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1535"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1455"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=532"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2328"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2249"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=532"
         Splits(0)._ColumnProps(22)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2090"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2011"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=532"
         Splits(0)._ColumnProps(28)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=1799"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1720"
         Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=532"
         Splits(0)._ColumnProps(34)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=532"
         Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=873"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=794"
         Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=530"
         Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(47)=   "Column(8).Width=1349"
         Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1270"
         Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=530"
         Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(52)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=532"
         Splits(0)._ColumnProps(56)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(58)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=532"
         Splits(0)._ColumnProps(62)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(64)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(67)=   "Column(11)._ColStyle=532"
         Splits(0)._ColumnProps(68)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(69)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(70)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(71)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(72)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(73)=   "Column(12)._ColStyle=532"
         Splits(0)._ColumnProps(74)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(75)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(76)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(77)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(78)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(79)=   "Column(13)._ColStyle=532"
         Splits(0)._ColumnProps(80)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(81)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(82)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(83)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(84)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(85)=   "Column(14)._ColStyle=532"
         Splits(0)._ColumnProps(86)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(87)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(88)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(89)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(90)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(91)=   "Column(15)._ColStyle=532"
         Splits(0)._ColumnProps(92)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(93)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(94)=   "Column(16).Width=1931"
         Splits(0)._ColumnProps(95)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(96)=   "Column(16)._WidthInPix=1852"
         Splits(0)._ColumnProps(97)=   "Column(16)._ColStyle=530"
         Splits(0)._ColumnProps(98)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(99)=   "Column(17).Width=1931"
         Splits(0)._ColumnProps(100)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(101)=   "Column(17)._WidthInPix=1852"
         Splits(0)._ColumnProps(102)=   "Column(17)._ColStyle=530"
         Splits(0)._ColumnProps(103)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(104)=   "Column(18).Width=2381"
         Splits(0)._ColumnProps(105)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(106)=   "Column(18)._WidthInPix=2302"
         Splits(0)._ColumnProps(107)=   "Column(18)._ColStyle=530"
         Splits(0)._ColumnProps(108)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(109)=   "Column(19).Width=1984"
         Splits(0)._ColumnProps(110)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(111)=   "Column(19)._WidthInPix=1905"
         Splits(0)._ColumnProps(112)=   "Column(19)._ColStyle=530"
         Splits(0)._ColumnProps(113)=   "Column(19).AllowFocus=0"
         Splits(0)._ColumnProps(114)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(115)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(116)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(117)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(118)=   "Column(20)._ColStyle=530"
         Splits(0)._ColumnProps(119)=   "Column(20).AllowFocus=0"
         Splits(0)._ColumnProps(120)=   "Column(20).Order=21"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTips        =   2
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HEFD8C2&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Arial"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HFCCFAB&,.bold=0"
         _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=Arial"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=94,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=90,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=87,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=88,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=89,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=98,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=95,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=96,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=97,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=66,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=70,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=74,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=78,.parent=13"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=75,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=76,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=77,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=82,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=79,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=80,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=81,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=86,.parent=13"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=83,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=84,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=85,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=102,.parent=13,.alignment=1"
         _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=106,.parent=13,.alignment=1"
         _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
         _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
         _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
         _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=110,.parent=13,.alignment=1"
         _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
         _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
         _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
         _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=114,.parent=13,.alignment=1"
         _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
         _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
         _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
         _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=118,.parent=13,.alignment=1"
         _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
         _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
         _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
         _StyleDefs(121) =   "Named:id=33:Normal"
         _StyleDefs(122) =   ":id=33,.parent=0,.bgcolor=&H80000005&"
         _StyleDefs(123) =   "Named:id=34:Heading"
         _StyleDefs(124) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(125) =   ":id=34,.wraptext=-1"
         _StyleDefs(126) =   "Named:id=35:Footing"
         _StyleDefs(127) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(128) =   "Named:id=36:Selected"
         _StyleDefs(129) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(130) =   "Named:id=37:Caption"
         _StyleDefs(131) =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(132) =   "Named:id=38:HighlightRow"
         _StyleDefs(133) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(134) =   "Named:id=39:EvenRow"
         _StyleDefs(135) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(136) =   "Named:id=40:OddRow"
         _StyleDefs(137) =   ":id=40,.parent=33"
         _StyleDefs(138) =   "Named:id=41:RecordSelector"
         _StyleDefs(139) =   ":id=41,.parent=34"
         _StyleDefs(140) =   "Named:id=42:FilterBar"
         _StyleDefs(141) =   ":id=42,.parent=33"
         _StyleDefs(142) =   "Named:id=0:"
         _StyleDefs(143) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(144) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(145) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(146) =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(147) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(148) =   ":id=0,.fontname=MS Sans Serif"
      End
      Begin TDBNumber6Ctl.TDBNumber tdbtPorc 
         Height          =   375
         Left            =   5490
         TabIndex        =   7
         Top             =   135
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "FrmManRegActRenFin.frx":145D
         Caption         =   "FrmManRegActRenFin.frx":147D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegActRenFin.frx":14E1
         Keys            =   "FrmManRegActRenFin.frx":14FF
         Spin            =   "FrmManRegActRenFin.frx":1539
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0.00"
         EditMode        =   2
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
         HighlightText   =   1
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
         ShowContextMenu =   1
         ValueVT         =   1703937
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label7 
         Caption         =   "PERIODO DE TRABAJO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   795
         TabIndex        =   17
         Top             =   255
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   705
         Width           =   915
      End
      Begin MSForms.CommandButton cmdGrabar 
         Height          =   375
         Left            =   3915
         TabIndex        =   15
         ToolTipText     =   "Grabar modificaciones"
         Top             =   540
         Width           =   1380
         Caption         =   " Grabar"
         PicturePosition =   327683
         Size            =   "2434;661"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   375
         Left            =   5475
         TabIndex        =   14
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   540
         Width           =   1380
         Caption         =   " Listar"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "FrmManRegActRenFin.frx":1561
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   11040
         TabIndex        =   13
         Top             =   540
         Width           =   1350
         Caption         =   "   Salir"
         PicturePosition =   327683
         Size            =   "2381;688"
         Picture         =   "FrmManRegActRenFin.frx":1AFB
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminarTodo 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
         Top             =   2955
         Visible         =   0   'False
         Width           =   1380
         Caption         =   " Eliminar Todo"
         PicturePosition =   327683
         Size            =   "2434;661"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminaItem 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Eliminar el movimientos seleccionado"
         Top             =   2490
         Visible         =   0   'False
         Width           =   1380
         Caption         =   " Eliminar Item"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "FrmManRegActRenFin.frx":2095
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Grabar modificaciones"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1380
         Caption         =   " Grabar"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "FrmManRegActRenFin.frx":262F
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   1080
         Visible         =   0   'False
         Width           =   1380
         Caption         =   " Listar"
         PicturePosition =   327683
         Size            =   "2434;661"
         Picture         =   "FrmManRegActRenFin.frx":2BC9
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsertarItem 
         Height          =   375
         Left            =   135
         TabIndex        =   10
         ToolTipText     =   "Insertar el movimientos seleccionado"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1380
         Caption         =   " Insertar Mov."
         PicturePosition =   327683
         Size            =   "2434;661"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   465
      Top             =   6510
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
            Picture         =   "FrmManRegActRenFin.frx":3163
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":353D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":3917
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":3CF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":40CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":44A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":487F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":4C59
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   6870
      Top             =   105
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
            Picture         =   "FrmManRegActRenFin.frx":5C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":620D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":67A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":6D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":72DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":7875
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":7E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":83A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":8943
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   6375
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
            Picture         =   "FrmManRegActRenFin.frx":8EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":9037
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":9191
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":92EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":9445
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":959F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":96F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":9853
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManRegActRenFin.frx":99AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc rsArreglo 
      Height          =   750
      Left            =   2535
      Top             =   7065
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LblCuenta 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3750
      TabIndex        =   18
      Top             =   390
      Width           =   75
   End
End
Attribute VB_Name = "FrmManRegActRenFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoCmdActivo As ADODB.Command
Dim arrDet As New XArrayDB
Dim l_item As Byte
Dim i As Integer
Dim Fila_ As Integer
Dim estdo_g As Boolean
'Dim AdoRsActivo As ADODB.Recordset
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdGrabar_Click()
On Error GoTo MIERROR
    Dim i As Integer
    Dim u As Integer
    For u = 0 To Me.tdbgActivo.ApproxCount
    
    
    
    
    Next
    
    For i = 0 To arrDet.Count(1) - 1 ' Me.rsArreglo.Recordset.RecordCount - 1
        If Existe(arrDet(i, 0), arrDet(i, 1)) Then
            Grabar "ACTUALIZAR", arrDet(i, 0), Me.tdbcMes.BoundText, arrDet(i, 1), _
            arrDet(i, 2), arrDet(i, 7), arrDet(i, 16), arrDet(i, 17), arrDet(i, 18), arrDet(i, 3)
        Else
            Grabar "GRABAR", arrDet(i, 0), Me.tdbcMes.BoundText, arrDet(i, 1), arrDet(i, 2), _
            arrDet(i, 7), arrDet(i, 16), arrDet(i, 17), arrDet(i, 18), arrDet(i, 3)
        End If
    Next
    MsgBox "Los Datos se grabaron con exito", 64
    Exit Sub
MIERROR:
    MsgBox Err.Description
End Sub

Private Sub cmdInsertarItem_Click()
'tdbgActivo.Row = l_item
'tdbgActivo.Col = 4
'
'    MsgBox (tdbgActivo.Text)
'
'    l_item = l_item + 1
    


    
    If estdo_g = False Then
    Fila_ = Fila_ + 1
    arrDet.ReDim 0, Fila_, 0, 20
    Set Me.tdbgActivo.Array = arrDet
    Me.tdbgActivo.ReBind
    estdo_g = True
    Else
        If MsgBox("Necesita Grabar el Item antes de Insertar otro, Desea Grabarlo", vbYesNo, "Ecbcont") = vbYes Then
             tdbgActivo.Row = Fila_
             
             MsgBox (tdbgActivo.Columns(1).Text)
             MsgBox (tdbgActivo.Columns(1).Value)
             
'             Grabar "GRABAR", arrDet(i, 0), Me.tdbcMes.BoundText, arrDet(i, 1), arrDet(i, 2), _
'             arrDet(i, 7), arrDet(i, 16), arrDet(i, 17), arrDet(i, 18), arrDet(i, 3)
             
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    Call CargarDatos
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Dim Mes As String
    
    Call Centrar_form(Me)

    Call LlenaCombos
'    Call LlenaComboMesApeAddItem(tdbcMes)
    
    DoEvents
    tdbcMes.ReBind
    
    If entro = False Then
        If gsPeriodo = "" Then
            tdbcMes.BoundText = "00"
        Else
            tdbcMes.BoundText = gsPeriodo
        End If
    Else
        If gsPeriodo >= "00" And gsPeriodo < "13" Then
            tdbcMes.BoundText = gsPeriodo
        Else
            tdbcMes.BoundText = "01"
        End If
    
    End If
    
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Dim entro As Boolean
    
    entro = False
    
    'If InStr(1, TituloSunat, "8.1") > 0 Or InStr(1, TituloSunat, "14.1") > 0 Then
        Call LlenaComboMesApeAddItem(tdbcMes)
        
        entro = True
    'Else
    '    Call LlenaComboMesApeAddItem(tdbcMes)
    'End If
    
    
    DoEvents
    tdbcMes.ReBind
    
    If entro = False Then
        If gsPeriodo = "" Then
            tdbcMes.BoundText = "00"
        Else
            tdbcMes.BoundText = gsPeriodo
        End If
    Else
        If gsPeriodo >= "00" And gsPeriodo < "13" Then
            tdbcMes.BoundText = gsPeriodo
        Else
            tdbcMes.BoundText = "01"
        End If
    
    End If
    
    
End Sub

Sub CargarDatos()
On Error GoTo MIERROR
    arrDet.ReDim 0, 0, 0, 0
    Set tdbgActivo.Array = arrDet
    ' *** Cargar el Almacen por Defecto
    l_item = 0
    i = 0
    Dim sSql As String

    sSql = "EXEC spCn_ActivosArrendamientoFinanciero 'LISTAR','" + gsEmpresa + "','" + _
    gsAnio + "','" + tdbcMes.BoundText + "'"
    rsArreglo.ConnectionString = gsCadenaConexion
    rsArreglo.RecordSource = sSql
    rsArreglo.Refresh
    arrDet.Clear
    estdo_g = False
    If rsArreglo.Recordset.RecordCount = 0 Then
    arrDet.ReDim 0, -1, 0, 20
    Fila_ = -1 + Fila_
    Else
    arrDet.ReDim 0, rsArreglo.Recordset.RecordCount - 1, 0, 20
    Fila_ = rsArreglo.Recordset.RecordCount - 1 + Fila_
    End If
    
    
    
       

    With rsArreglo.Recordset
        Do While Not .EOF
            i = i + 1
            arrDet(i - 1, 0) = .Fields("Ase_cNummov").Value
            arrDet(i - 1, 1) = .Fields("Ase_nVoucher").Value
            arrDet(i - 1, 2) = .Fields("Act_dFechaContrato").Value
            arrDet(i - 1, 3) = .Fields("Pla_cNombreCuenta").Value
            arrDet(i - 1, 4) = .Fields("Mon_cNombreLargo").Value
            arrDet(i - 1, 5) = .Fields("Asd_dFecDoc").Value
            arrDet(i - 1, 6) = .Fields("Pla_cCuentaContable").Value
            arrDet(i - 1, 7) = .Fields("Act_nVidaUtil").Value
            arrDet(i - 1, 8) = FormatDec(.Fields("Asd_nDebeSoles").Value)
            arrDet(i - 1, 9) = FormatDec(.Fields("Asd_nHaberSoles").Value)
            arrDet(i - 1, 10) = .Fields("Ten_cTipoEntidad").Value
            arrDet(i - 1, 11) = .Fields("Ent_cCodEntidad").Value
            arrDet(i - 1, 12) = .Fields("Asd_cTipoDoc").Value
            arrDet(i - 1, 13) = .Fields("Asd_cSerieDoc").Value
            arrDet(i - 1, 14) = .Fields("Asd_cNumDoc").Value
            arrDet(i - 1, 15) = .Fields("Asd_cTipoMoneda").Value
            arrDet(i - 1, 16) = FormatDec(.Fields("Act_nDepreciacionPorc").Value)
            arrDet(i - 1, 17) = FormatDec(.Fields("Act_nDepreciacionAcumulada").Value)
            arrDet(i - 1, 18) = FormatDec(.Fields("Act_nDepreciacionEjercicio").Value)
            
            ''arrDet(i - 1, 19) = FormatDec(0)
            'arrDet(i - 1, 20) = FormatDec(0)
            
        '.Columns(18).Value = FormatDec(.Fields("Asd_nDebeSoles").Value * (.Fields("Act_nDepreciacionPorc").Value / 100))
            arrDet(i - 1, 19) = FormatDec(.Fields("Act_Deprecion").Value)
            arrDet(i - 1, 20) = FormatDec(.Fields("Act_totales").Value)
            
            'Call Calcular
            .MoveNext
        Loop
    End With
    Set Me.tdbgActivo.Array = arrDet
    
    Me.tdbgActivo.ReBind
    l_item = i + 1
    tdbgActivo.Splits(0).Locked = False
    Exit Sub
MIERROR:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsArreglo.Recordset.Close
    Set arrDet = Nothing
End Sub

Private Sub tdbcMes_ItemChange()
    Call CargarDatos
End Sub

Private Sub tdbgActivo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo MIERROR
    With Me.tdbgActivo
        Select Case LastCol
            Case 2
                arrDet(.Bookmark, 2) = dtpFechaDoc
            Case 3
                arrDet(.Bookmark, 3) = .Columns(3).Value
            Case 16, 17
                Call Calcular
            End Select
    End With
    Exit Sub
MIERROR:
End Sub

Private Function Grabar(ByRef Tipo As String, ByRef Ase_cNummov As String, _
ByRef Per_cPeriodo As String, ByRef Ase_nVoucher As String, ByRef Act_dFechaContrato As String, _
ByRef Act_nVidaUtil As Integer, ByRef Act_nDepreciacionPorc As Double, _
ByRef Act_nDepreciacionAcumulada As Double, ByRef Act_nDepreciacionEjercicio As Double, _
Act_cTipoBien As String) As Boolean
On Error GoTo MIERROR
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "spCn_ActivosArrendamientoFinanciero"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Tipo", adVarChar, adParamInput, 20, Tipo)
        .Parameters.Append .CreateParameter("@Emp_cCodigo", adChar, adParamInput, 3, gsEmpresa)
        .Parameters.Append .CreateParameter("@Pan_cAnio", adChar, adParamInput, 4, gsAnio)
        .Parameters.Append .CreateParameter("@Per_cPeriodo", adChar, adParamInput, 2, Per_cPeriodo)
        .Parameters.Append .CreateParameter("@Ase_cNummov", adChar, adParamInput, 10, Ase_cNummov)
        .Parameters.Append .CreateParameter("@Ase_nVoucher", adChar, adParamInput, 10, Ase_nVoucher)
        .Parameters.Append .CreateParameter("@Act_dFechaContrato", adVarChar, adParamInput, 10, Act_dFechaContrato)
        .Parameters.Append .CreateParameter("@Act_nVidaUtil", adInteger, adParamInput, , Act_nVidaUtil)
        .Parameters.Append .CreateParameter("@Act_nDepreciacionPorc", adDouble, adParamInput, , Act_nDepreciacionPorc)
        .Parameters.Append .CreateParameter("@Act_nDepreciacionAcumulada", adDouble, adParamInput, , Act_nDepreciacionAcumulada)
        .Parameters.Append .CreateParameter("@Act_nDepreciacionEjercicio", adDouble, adParamInput, , Act_nDepreciacionEjercicio)
        .Parameters.Append .CreateParameter("@Act_cTipoBien", adVarChar, adParamInput, 120, Act_cTipoBien)

        .Execute
    End With
    
    Set VarCmd = Nothing
    Grabar = True
    Exit Function
MIERROR:
    MsgBox Err.Description
'    RESUME
Grabar = False
    Set VarCmd = Nothing
End Function

Private Function Existe(ByRef Ase_cNummov As String, ByRef Ase_nVoucher As String) As Boolean
On Error GoTo MIERROR
    Dim AdoRsExiste As ADODB.Recordset
    Set AdoRsExiste = New ADODB.Recordset
    Dim sSql As String
    Existe = False
    sSql = "spCn_ActivosArrendamientoFinanciero 'EXISTE','" & gsEmpresa & "','" & gsAnio & "','" & _
    tdbcMes.BoundText & "','" & Ase_cNummov & "','" & Ase_nVoucher & "','',1,0,0,0"
    rsArreglo.RecordSource = sSql
    rsArreglo.Refresh
    
    If rsArreglo.Recordset.RecordCount > 0 Then
        Existe = True
    End If
    Set AdoRsExiste = Nothing
    Exit Function
MIERROR:
    MsgBox Err.Description
End Function

Private Sub Calcular()
Dim Fila As Integer
On Error GoTo MIERROR
    With Me.tdbgActivo
    Fila = IIf(IsNull(.Bookmark), 1, .Bookmark)
    If Fila = 0 Then
        .Columns(18).Value = FormatDec(.Columns(8).Value * (.Columns(16).Value / 100))
        .Columns(19).Value = FormatDec(.Columns(17).Value + .Columns(18).Value)
        .Columns(20).Value = FormatDec(.Columns(8).Value - .Columns(19).Value)
    Else
'        arrDet(i - 1, 18) = FormatDec(arrDet(i - 1, 8) * (arrDet(i - 1, 16) / 100))
'        arrDet(i - 1, 19) = FormatDec(arrDet(i - 1, 17) + arrDet(i - 1, 18))
'        arrDet(i - 1, 20) = FormatDec(arrDet(i - 1, 8) - arrDet(i - 1, 19))
        
'        arrDet(i, 18) = FormatDec(arrDet(i, 8) * (arrDet(i, 16) / 100))
'        arrDet(i, 19) = FormatDec(arrDet(i, 17) + arrDet(i, 18))
'        arrDet(i, 20) = FormatDec(arrDet(i, 8) - arrDet(i, 19))

                    arrDet(i - 1, 16) = .Columns(16).Value
                    arrDet(i - 1, 17) = .Columns(17).Value
                    arrDet(i - 1, 18) = FormatDec(.Columns(8).Value * (.Columns(16).Value / 100)) 'FormatDec(arrDet(i - 1, 8) * (arrDet(i - 1, 16) / 100))
                    arrDet(i - 1, 19) = FormatDec(.Columns(17).Value + arrDet(i - 1, 18))
                    arrDet(i - 1, 20) = FormatDec(.Columns(8).Value - arrDet(i - 1, 19))
            
            
                    .Columns(18).Value = arrDet(i - 1, 18)
                    .Columns(19).Value = arrDet(i - 1, 19)
                    .Columns(20).Value = arrDet(i - 1, 20)
'        .Columns(18).Value = FormatDec(.Columns(8).Value * (.Columns(16).Value / 100))
'        .Columns(19).Value = FormatDec(.Columns(17).Value + .Columns(18).Value)
'        .Columns(20).Value = FormatDec(.Columns(8).Value - .Columns(19).Value)
    End If
    End With
    Exit Sub
MIERROR:
End Sub
