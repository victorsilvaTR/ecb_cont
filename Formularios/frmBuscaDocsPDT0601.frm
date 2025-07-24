VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBuscaDocsPDT0601 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Documentos PDT0601"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17340
   Icon            =   "frmBuscaDocsPDT0601.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   17340
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7950
      Begin TDBText6Ctl.TDBText tdbtNombreEntidad 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   270
         Width           =   4620
         _Version        =   65536
         _ExtentX        =   8149
         _ExtentY        =   556
         Caption         =   "frmBuscaDocsPDT0601.frx":0ECA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBuscaDocsPDT0601.frx":0F36
         Key             =   "frmBuscaDocsPDT0601.frx":0F54
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
         Format          =   "A"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
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
      Begin TDBText6Ctl.TDBText tdbtSerie 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   990
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   556
         Caption         =   "frmBuscaDocsPDT0601.frx":0F96
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBuscaDocsPDT0601.frx":1002
         Key             =   "frmBuscaDocsPDT0601.frx":1020
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
         Format          =   "A9"
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
      Begin TDBText6Ctl.TDBText tdbtNumero 
         Height          =   315
         Left            =   4275
         TabIndex        =   3
         Top             =   990
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   556
         Caption         =   "frmBuscaDocsPDT0601.frx":1062
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBuscaDocsPDT0601.frx":10CE
         Key             =   "frmBuscaDocsPDT0601.frx":10EC
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
         Format          =   "A9"
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
      Begin TDBText6Ctl.TDBText tdbtRuc 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   630
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "frmBuscaDocsPDT0601.frx":1120
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBuscaDocsPDT0601.frx":118C
         Key             =   "frmBuscaDocsPDT0601.frx":11AA
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
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ruc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   1035
         Width           =   450
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgPDT0601 
      Height          =   4755
      Left            =   45
      TabIndex        =   10
      Top             =   1665
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   8387
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "R.U.C."
      Columns(0).DataField=   "Ent_nRuc"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Entidad"
      Columns(1).DataField=   "Ent_cPersona"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cod.Tipo de Documento"
      Columns(2).DataField=   "Ent_cTipoDoc2"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "T.D."
      Columns(3).DataField=   "Dsc_Ent_cTipoDoc2"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Serie"
      Columns(4).DataField=   "Asd_cSerieDoc"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Número"
      Columns(5).DataField=   "Asd_cNumDoc"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fecha Emisión"
      Columns(6).DataField=   "Fecha"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fecha Pago"
      Columns(7).DataField=   "FechaCanc"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Importe"
      Columns(8).DataField=   "Asd_nHaberSoles"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Cod. Tipo Comp. Emi."
      Columns(9).DataField=   "Cod_Asd_cTipoDoc2"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Tipo Comp. Emitido"
      Columns(10).DataField=   "Asd_cTipoDoc2"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2434"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2355"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=13150"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=13070"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=3784"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3704"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1561"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1482"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1931"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1852"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2487"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2408"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=2355"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2090"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2011"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=1879"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1799"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=3545"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=3466"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=513"
      Splits(0)._ColumnProps(61)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(63)=   "Column(10).Width=1958"
      Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=1879"
      Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(69)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(74)=   "Column(11).Order=12"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=16,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H80000007&"
      _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000E&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Arial"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=Arial"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=122,.parent=13,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=119,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=120,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=121,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=118,.parent=13,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=138,.parent=13,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=135,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=136,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=137,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=134,.parent=13,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=131,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=132,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=133,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=130,.parent=13,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=126,.parent=13,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=123,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=124,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=125,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
      _StyleDefs(85)  =   "Named:id=33:Normal"
      _StyleDefs(86)  =   ":id=33,.parent=0"
      _StyleDefs(87)  =   "Named:id=34:Heading"
      _StyleDefs(88)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   ":id=34,.wraptext=-1"
      _StyleDefs(90)  =   "Named:id=35:Footing"
      _StyleDefs(91)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   "Named:id=36:Selected"
      _StyleDefs(93)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(94)  =   "Named:id=37:Caption"
      _StyleDefs(95)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(96)  =   "Named:id=38:HighlightRow"
      _StyleDefs(97)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(98)  =   "Named:id=39:EvenRow"
      _StyleDefs(99)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(100) =   "Named:id=40:OddRow"
      _StyleDefs(101) =   ":id=40,.parent=33"
      _StyleDefs(102) =   "Named:id=41:RecordSelector"
      _StyleDefs(103) =   ":id=41,.parent=34"
      _StyleDefs(104) =   "Named:id=42:FilterBar"
      _StyleDefs(105) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Seleccione un documento haciendo doble click a un elemento de la lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8820
      TabIndex        =   9
      Top             =   1125
      Width           =   8415
   End
End
Attribute VB_Name = "frmBuscaDocsPDT0601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lrsPDT0601 As ADODB.Recordset
Dim lArrDatos As New XArrayDB

Dim nCol_cRuc As Integer
Dim nCol_cEntidad As Integer
Dim nCol_cTipDocu As Integer
Dim nCol_cDscTipoDocu As Integer
Dim nCol_cSerDocu As Integer
Dim nCol_cNroDoc As Integer
Dim nCol_cFechaEmision As Integer
Dim nCol_cFechaPago As Integer
Dim nCol_cImporte As Integer
Dim nCol_cCodTipoCompEmi As Integer
Dim nCol_cTipoCompEmi As Integer
Dim nCol_cRetencion As Integer

Dim NUM_FILAS As Integer
Dim NUM_COLUMNAS As Integer
Private Sub Form_Load()
    
 Call IniciaVariables
 
 NUM_FILAS = 0
 NUM_COLUMNAS = 12
 
 Call Centrar_form(Me)
    
 Call LlenaPDT0601
 Call ConfigurarColumnas
 
End Sub
Private Sub Form_Resize()
     If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        tdbgPDT0601.Width = Me.Width - 200
        tdbgPDT0601.Height = Me.Height - 1600
    End If
    
    Exit Sub
serror:
 Resume Next
End Sub

Public Sub Form_Unload(Cancel As Integer)

 Call CerrarRecordSet(lrsPDT0601)
 Set frmBuscaDocsPDT0601 = Nothing
 Unload Me
 
End Sub
Private Sub LlenaPDT0601()
On Error GoTo Control

 Dim i As Integer
 Dim Col As Integer
 
 Dim sqlSp As String
 Dim clDatos As clsMantoTablas
 Dim arrDatos() As Variant
    
 Set clDatos = New clsMantoTablas
 Set lrsPDT0601 = New ADODB.Recordset
 Set tdbgPDT0601.DataSource = Nothing
 
 sqlSp = "spCn_GrabaPDT0601 'BUSCAR_ITEM','" & gsEmpresa & "','" & gsAnio & "', '" & frmManPDT0601.tdbcMes.BoundText & "', '" & IIf(frmManPDT0601.tdbcLibro.BoundText = "00", "", frmManPDT0601.tdbcLibro.BoundText) & "'"
   
 arrDatos = Array(sqlSp)
 Set lrsPDT0601 = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
 
If Not lrsPDT0601 Is Nothing Then
 Set tdbgPDT0601.DataSource = lrsPDT0601
 Exit Sub
End If

 tdbgPDT0601.ReBind
Exit Sub
Control:
 MsgBox Err.Description
End Sub
Private Sub ConfigurarColumnas()
 Call OcultaColumna(nCol_cTipDocu)
 Call OcultaColumna(nCol_cCodTipoCompEmi)
End Sub

Sub IniciaVariables()
 nCol_cRuc = 0
 nCol_cEntidad = 1
 nCol_cTipDocu = 2
 nCol_cDscTipoDocu = 3
 nCol_cSerDocu = 4
 nCol_cNroDoc = 5
 nCol_cFechaEmision = 6
 nCol_cFechaPago = 7
 nCol_cImporte = 8
 nCol_cCodTipoCompEmi = 9
 nCol_cTipoCompEmi = 10
 nCol_cRetencion = 11
End Sub

Private Sub OcultaColumna(nCol As Integer)
 On Error GoTo serror
    With tdbgPDT0601
        .Columns(nCol).Visible = False
        .Columns(nCol).Width = 0
        .Splits(0).Columns(nCol).AllowFocus = False
        .Splits(0).Columns(nCol).AllowSizing = False
    End With
    Exit Sub
serror:
 MsgBox Err.Description
End Sub
Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    
    Dim filtros(5) As String
    Dim i As Integer
    cadena = ""
    If CE(tdbtNombreEntidad) <> "" Then filtros(0) = "Ent_cPersona like '*" & tdbtNombreEntidad & "*'"
    If CE(tdbtSerie) <> "" Then filtros(1) = "Asd_cSerieDoc like '*" & tdbtSerie & "*'"
    If CE(tdbtNumero) <> "" Then filtros(2) = "Asd_cNumDoc like '*" & tdbtNumero & "*'"
    If CE(tdbtRuc) <> "" Then filtros(3) = "Ent_nRuc like '" & tdbtRuc & "*'"
    For i = 0 To 5
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    If lrsPDT0601 Is Nothing = False Then lrsPDT0601.Filter = 0
    ' *** Filtrando segun campos
    If Not lrsPDT0601 Is Nothing Then
        If Not (lrsPDT0601.BOF And lrsPDT0601.EOF) Then
            If CE(cadena) <> "" Then
                lrsPDT0601.Filter = cadena
            Else
                lrsPDT0601.Filter = 0
            End If
        End If
    End If
    

End Sub
Private Sub tdbgPDT0601_DblClick()
 Call frmManPDT0601.LlenaDatos
 frmManPDT0601.InsItem = True
End Sub

Private Sub tdbtNombreEntidad_Change()
    If gsKey = 219 Then
        tdbtNombreEntidad = Replace(tdbtNombreEntidad, "'", "")
        tdbtNombreEntidad.SelStart = Len(tdbtNombreEntidad)
    End If
    
    Call FiltrarRecordSet

End Sub
Private Sub tdbtNombreEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
 gsKey = KeyCode
 If KeyCode = 40 Then
  Siguiente
  KeyCode = 0
 End If
 If KeyCode = 38 Then
  Anterior
  KeyCode = 0
 End If
End Sub
Private Sub tdbtNumero_Change()
 Call FiltrarRecordSet
End Sub
Private Sub tdbtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 40 Then
  Siguiente
  KeyCode = 0
 End If
 If KeyCode = 38 Then
  Anterior
  KeyCode = 0
 End If
End Sub

Private Sub tdbtRuc_Change()
 Call FiltrarRecordSet
End Sub
Private Sub tdbtSerie_Change()
 Call FiltrarRecordSet
End Sub
Private Sub tdbtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 40 Then
  Siguiente
  KeyCode = 0
 End If
 If KeyCode = 38 Then
  Anterior
  KeyCode = 0
 End If
End Sub
Private Sub Siguiente()
 tdbgPDT0601.MoveNext
 If tdbgPDT0601.EOF Then tdbgPDT0601.MoveLast
End Sub
Private Sub Anterior()
 tdbgPDT0601.MovePrevious
 If tdbgPDT0601.BOF Then tdbgPDT0601.MoveFirst
End Sub
