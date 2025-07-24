VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigOperaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Operaciones"
   ClientHeight    =   6075
   ClientLeft      =   1530
   ClientTop       =   2550
   ClientWidth     =   8355
   Icon            =   "frmConfigOperaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   8355
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Configuración de Operaciones"
      TabPicture(0)   =   "frmConfigOperaciones.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tbrOpciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imglstTool"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imglstdisabled"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "sstParmetros"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Configuración de Libro vs Cuenta"
      TabPicture(1)   =   "frmConfigOperaciones.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tbrOpcioneslibrocuenta"
      Tab(1).Control(1)=   "sstLibros"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         Caption         =   "Cuentas Contables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   7965
         Begin TrueOleDBGrid70.TDBGrid tdbgCuentas 
            Height          =   1800
            Left            =   255
            TabIndex        =   14
            Top             =   270
            Width           =   7420
            _ExtentX        =   13097
            _ExtentY        =   3175
            _LayoutType     =   4
            _RowHeight      =   18
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "Pla_cCuentaContable"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "Pla_cNombreCuenta"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=10901"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10821"
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
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=196,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parametros de Operaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   8100
         Begin TrueOleDBGrid70.TDBGrid tdbgParametros 
            Height          =   1815
            Left            =   60
            TabIndex        =   12
            Top             =   270
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   3201
            _LayoutType     =   4
            _RowHeight      =   18
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Empresa"
            Columns(0).DataField=   "Emp_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "Cop_cCodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "Cop_cDescripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo"
            Columns(3).DataField=   "Cop_cTipo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "Cop_cEstado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=10874"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=10795"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(29)=   "Column(4).Width=8176"
            Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=8096"
            Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=196,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(57)  =   "Named:id=33:Normal"
            _StyleDefs(58)  =   ":id=33,.parent=0"
            _StyleDefs(59)  =   "Named:id=34:Heading"
            _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(61)  =   ":id=34,.wraptext=-1"
            _StyleDefs(62)  =   "Named:id=35:Footing"
            _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   "Named:id=36:Selected"
            _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(66)  =   "Named:id=37:Caption"
            _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(68)  =   "Named:id=38:HighlightRow"
            _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(70)  =   "Named:id=39:EvenRow"
            _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(72)  =   "Named:id=40:OddRow"
            _StyleDefs(73)  =   ":id=40,.parent=33"
            _StyleDefs(74)  =   "Named:id=41:RecordSelector"
            _StyleDefs(75)  =   ":id=41,.parent=34"
            _StyleDefs(76)  =   "Named:id=42:FilterBar"
            _StyleDefs(77)  =   ":id=42,.parent=33"
         End
      End
      Begin TabDlg.SSTab sstParmetros 
         Height          =   2760
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   4868
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Lista de Valores de Parametros"
         TabPicture(0)   =   "frmConfigOperaciones.frx":0F02
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Edición de Datos"
         TabPicture(1)   =   "frmConfigOperaciones.frx":0F1E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            Caption         =   "Valor de Parametro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Left            =   60
            TabIndex        =   9
            Top             =   465
            Width           =   7935
            Begin TrueOleDBGrid70.TDBGrid tdbgValorParam 
               Height          =   1500
               Left            =   60
               TabIndex        =   10
               Top             =   330
               Width           =   7755
               _ExtentX        =   13679
               _ExtentY        =   2646
               _LayoutType     =   4
               _RowHeight      =   18
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Empresa"
               Columns(0).DataField=   "Emp_cCodigo"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Año"
               Columns(1).DataField=   "Pan_cAnio"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Código"
               Columns(2).DataField=   "Cop_cCodigo"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Valor"
               Columns(3).DataField=   "Cod_cValorParam"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Descripción"
               Columns(4).DataField=   "Descripcion"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "IGV %"
               Columns(5).DataField=   "Cod_nIgvPorc"
               Columns(5).NumberFormat=   "Standard"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "Estado"
               Columns(6).DataField=   "Cod_cEstado"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   7
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).Locked=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=7"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=532"
               Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(9)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
               Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=532"
               Splits(0)._ColumnProps(15)=   "Column(1).Visible=0"
               Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(17)=   "Column(2).Width=1164"
               Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1085"
               Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=532"
               Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
               Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(24)=   "Column(3).Width=1746"
               Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1667"
               Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=532"
               Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(30)=   "Column(4).Width=9260"
               Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=9181"
               Splits(0)._ColumnProps(33)=   "Column(4)._EditAlways=0"
               Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=532"
               Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(36)=   "Column(5).Width=1085"
               Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=1005"
               Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
               Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=532"
               Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(42)=   "Column(6).Width=1270"
               Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=1191"
               Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
               Splits(0)._ColumnProps(46)=   "Column(6).AllowSizing=0"
               Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=532"
               Splits(0)._ColumnProps(48)=   "Column(6).Visible=0"
               Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
               _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
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
               _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
               _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
               _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
               _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
               _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
               _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
               _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
               _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
               _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
               _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
               _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
               _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
               _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
               _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
               _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
               _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
               _StyleDefs(65)  =   "Named:id=33:Normal"
               _StyleDefs(66)  =   ":id=33,.parent=0"
               _StyleDefs(67)  =   "Named:id=34:Heading"
               _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(69)  =   ":id=34,.wraptext=-1"
               _StyleDefs(70)  =   "Named:id=35:Footing"
               _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(72)  =   "Named:id=36:Selected"
               _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(74)  =   "Named:id=37:Caption"
               _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(76)  =   "Named:id=38:HighlightRow"
               _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(78)  =   "Named:id=39:EvenRow"
               _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(80)  =   "Named:id=40:OddRow"
               _StyleDefs(81)  =   ":id=40,.parent=33"
               _StyleDefs(82)  =   "Named:id=41:RecordSelector"
               _StyleDefs(83)  =   ":id=41,.parent=34"
               _StyleDefs(84)  =   "Named:id=42:FilterBar"
               _StyleDefs(85)  =   ":id=42,.parent=33"
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1830
            Left            =   -74775
            TabIndex        =   2
            Top             =   540
            Width           =   7050
            Begin TDBNumber6Ctl.TDBNumber tdbnIgv 
               Height          =   315
               Left            =   1725
               TabIndex        =   3
               Tag             =   "Tag"
               Top             =   1275
               Visible         =   0   'False
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   556
               Calculator      =   "frmConfigOperaciones.frx":0F3A
               Caption         =   "frmConfigOperaciones.frx":0F5A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":0FC6
               Keys            =   "frmConfigOperaciones.frx":0FE4
               Spin            =   "frmConfigOperaciones.frx":102C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.00"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.00"
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
               ValueVT         =   1245189
               Value           =   0
               MaxValueVT      =   1768751109
               MinValueVT      =   1629487109
            End
            Begin TDBText6Ctl.TDBText tdbtDescripción 
               Height          =   315
               Left            =   3105
               TabIndex        =   4
               Tag             =   "Tag"
               Top             =   810
               Width           =   3675
               _Version        =   65536
               _ExtentX        =   6482
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":1054
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":10C0
               Key             =   "frmConfigOperaciones.frx":10DE
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
            Begin TDBText6Ctl.TDBText tdbtValorParam 
               Height          =   315
               Left            =   1710
               TabIndex        =   5
               Tag             =   "Tag"
               Top             =   810
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":1120
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":118C
               Key             =   "frmConfigOperaciones.frx":11AA
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
               Format          =   "9A"
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
            Begin VB.Label Label2 
               Caption         =   "Valor de Parametro"
               Height          =   225
               Left            =   285
               TabIndex        =   8
               Top             =   885
               Width           =   1470
            End
            Begin VB.Label lblMante 
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
               Height          =   450
               Left            =   225
               TabIndex        =   7
               Top             =   270
               Width           =   1605
            End
            Begin VB.Label lblIgv 
               Caption         =   "IGV %"
               Height          =   255
               Left            =   330
               TabIndex        =   6
               Top             =   1335
               Visible         =   0   'False
               Width           =   720
            End
         End
      End
      Begin TabDlg.SSTab sstLibros 
         Height          =   2700
         Left            =   -74880
         TabIndex        =   15
         Top             =   3135
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   4763
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Libros"
         TabPicture(0)   =   "frmConfigOperaciones.frx":11EC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Edición de Datos"
         TabPicture(1)   =   "frmConfigOperaciones.frx":1208
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame7"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   -74775
            TabIndex        =   23
            Top             =   495
            Width           =   7560
            Begin TDBText6Ctl.TDBText tdbtDescripcion 
               Height          =   315
               Left            =   2355
               TabIndex        =   24
               Tag             =   "Tag"
               Top             =   840
               Width           =   3510
               _Version        =   65536
               _ExtentX        =   6191
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":1224
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":1290
               Key             =   "frmConfigOperaciones.frx":12AE
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
            Begin TDBText6Ctl.TDBText tdbtLibro 
               Height          =   315
               Left            =   1590
               TabIndex        =   25
               Tag             =   "Tag"
               Top             =   840
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":12F0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":135C
               Key             =   "frmConfigOperaciones.frx":137A
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
               MaxLength       =   2
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
            Begin VB.Label lblMante2 
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
               Height          =   285
               Left            =   225
               TabIndex        =   27
               Top             =   330
               Width           =   3525
            End
            Begin VB.Label Label3 
               Caption         =   "Libro"
               Height          =   225
               Left            =   285
               TabIndex        =   26
               Top             =   885
               Width           =   1470
            End
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1680
            Left            =   -74640
            TabIndex        =   18
            Top             =   555
            Width           =   6120
            Begin TDBText6Ctl.TDBText TDBText1 
               Height          =   315
               Left            =   2355
               TabIndex        =   19
               Tag             =   "Tag"
               Top             =   840
               Width           =   3510
               _Version        =   65536
               _ExtentX        =   6191
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":13BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":1428
               Key             =   "frmConfigOperaciones.frx":1446
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
            Begin TDBText6Ctl.TDBText tdbtTipoDoc 
               Height          =   315
               Left            =   1590
               TabIndex        =   20
               Tag             =   "Tag"
               Top             =   840
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
               _ExtentY        =   556
               Caption         =   "frmConfigOperaciones.frx":1488
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmConfigOperaciones.frx":14F4
               Key             =   "frmConfigOperaciones.frx":1512
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
               MaxLength       =   2
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
               Height          =   285
               Left            =   225
               TabIndex        =   22
               Top             =   330
               Width           =   3525
            End
            Begin VB.Label Label7 
               Caption         =   "Tipo Documento"
               Height          =   225
               Left            =   285
               TabIndex        =   21
               Top             =   885
               Width           =   1470
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Tipo de Documentos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   300
            TabIndex        =   16
            Top             =   405
            Width           =   7380
            Begin TrueOleDBGrid70.TDBGrid tdbgLibro 
               Height          =   1725
               Left            =   150
               TabIndex        =   17
               Top             =   315
               Width           =   7065
               _ExtentX        =   12462
               _ExtentY        =   3043
               _LayoutType     =   4
               _RowHeight      =   18
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   "Lib_cTipoLibro"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Codigo"
               Columns(1).DataField=   "Lib_cTipoLibro"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Descripción"
               Columns(2).DataField=   "Lib_cDescripcion"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).Locked=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=291"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=212"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
               Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=1376"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1296"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=532"
               Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(14)=   "Column(2).Width=9763"
               Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=9684"
               Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=532"
               Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
               _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
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
               _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
               _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
               _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
               _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
               _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
               _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
               _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
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
         End
      End
      Begin MSComctlLib.ImageList imglstdisabled 
         Left            =   11160
         Top             =   2790
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
               Picture         =   "frmConfigOperaciones.frx":1554
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":16AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1808
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1962
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1ABC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1C16
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1D70
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":1ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":2024
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstTool 
         Left            =   11160
         Top             =   2115
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
               Picture         =   "frmConfigOperaciones.frx":217E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":2718
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":2CB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":324C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":37E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":3D80
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":431A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":48B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConfigOperaciones.frx":4E4E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrOpciones 
         Height          =   330
         Left            =   135
         TabIndex        =   28
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
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
      Begin MSComctlLib.Toolbar tbrOpcioneslibrocuenta 
         Height          =   330
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15
      Top             =   375
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
            Picture         =   "frmConfigOperaciones.frx":53E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":57C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":5B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":5F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":6350
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":672A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":6B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigOperaciones.frx":6EDE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConfigOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmConfigOperaciones
'    Project    : Contabilidad
'
'    Description: Formulario de configuracion de operaciones
'--------------------------------------------------------------------------------
Option Explicit
Dim lrsParam    As ADODB.Recordset
Dim lrsParam2   As ADODB.Recordset
Dim lrsValor    As ADODB.Recordset
Dim lrsLibro    As ADODB.Recordset
Dim lArrMnt()   As Variant
Dim lArrMnt2()  As Variant
Dim lTipoMnt    As String
Dim sqlSp       As String
Dim Control     As String
Dim lRegElim    As Boolean

Dim gsGrupo As String
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grupo
' Description:       Propiedad de asignacion de grupo
'
' Parameters :       Grupo (String)
'--------------------------------------------------------------------------------
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pLlenarParametros
' Description:       Procedimiento de llenado codigo de operaciones y cuentas contables
'
' Parameters :       cAdmin (Boolean = False)
'--------------------------------------------------------------------------------
Sub pLlenarParametros(Optional cAdmin As Boolean = False)
    Dim cValor As String
    cValor = IIf(cAdmin = True, "1", "0")
    
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim arrDatos2() As Variant
    Set clDatos = New clsMantoTablas
    sqlSp = "spCNT_CONFIG_OPERA 'BUSCARTODOS','" & gsEmpresa & "', null,null,'DOC',null,null,null,'" & cValor & "'"
    arrDatos = Array(sqlSp)
    Set lrsParam = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsParam Is Nothing Then Exit Sub
    tdbgParametros.DataSource = lrsParam
    
    DoEvents
    
    sqlSp = "spCn_RegLibroCta 'BUSCARTODOSCUENTAS','" & gsEmpresa & "','','','" & gsAnio & "'"
    arrDatos2 = Array(sqlSp)
    Set lrsParam2 = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos2())
    If lrsParam2 Is Nothing Then Exit Sub
    tdbgCuentas.DataSource = lrsParam2
   
    Set clDatos = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pLlenarValoresParametros
' Description:       Procedimiento de llenado del detalle de codigos de libros y valores de las configuracion de operaciones
'
' Parameters :       cAdmin (Boolean = False)
'--------------------------------------------------------------------------------
Sub pLlenarValoresParametros(Optional cAdmin As Boolean = False)
    Dim cValor As String
    cValor = IIf(cAdmin = True, "1", "0")

    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim arrDatos2() As Variant
    Set clDatos = New clsMantoTablas
    Set tdbgValorParam.DataSource = Nothing
    Set tdbgLibro.DataSource = Nothing
    
    sqlSp = "spCND_CONFIG_OPERA 'BUSCARTODOS','" & gsEmpresa & "','" & gsAnio & "', null,null,null,null,null,null,'" & cValor & "'"
    arrDatos = Array(sqlSp)
    Set lrsValor = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If lrsValor Is Nothing Then Exit Sub
    tdbgValorParam.DataSource = lrsValor
    
    sqlSp = "spCn_RegLibroCta 'BUSCARTODOSLIBRO','" & gsEmpresa & "','','','" & gsAnio & "'"
    arrDatos2 = Array(sqlSp)
    Set lrsLibro = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos2())
    If lrsLibro Is Nothing Then Exit Sub
    tdbgLibro.DataSource = lrsLibro
    
    Set clDatos = Nothing
End Sub


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
        Case 27:
            If Not sstParmetros.TabEnabled(1) Then
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar esta Operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                Else
                    pSetFocus tdbtValorParam
                End If
            End If
        Case 113:
            If SSTab1.Tab = 0 Then
                If tbrOpciones.Buttons(1).Enabled Then ManNuevo
            Else
                If tbrOpcioneslibrocuenta.Buttons(1).Enabled Then ManNuevo2
            End If
                    
        Case 115:
            If SSTab1.Tab = 0 Then
                If tbrOpciones.Buttons(3).Enabled Then Grabar
            Else
                If tbrOpcioneslibrocuenta.Buttons(3).Enabled Then Grabar2
            End If
        
        Case 116:
            If SSTab1.Tab = 0 Then
                If tbrOpciones.Buttons(4).Enabled Then Borrar
            Else
                If tbrOpcioneslibrocuenta.Buttons(4).Enabled Then Borrar2
            End If
        
        Case 117:
            If SSTab1.Tab = 0 Then
                If tbrOpciones.Buttons(5).Enabled Then Editar
            Else
                If tbrOpcioneslibrocuenta.Buttons(5).Enabled Then Editar2
            End If
        

        Case vbKeyF1:
                  If Shift = 1 Then
                     If InputBox("Ingrese la clave del Administrador" & Salto(1) & "Para ver toda la config. del sistema", "Seguridad", "") = "977611" Then
                        SSTab1.TabVisible(0) = True
                        Call CargaTodosLosDatosOP(True)
                     End If
                  Else
                     'SSTab1.TabVisible(0) = False
                     'Call CargaTodosLosDatosOP(False)
                  End If
        Case vbKeyF1:
                  If Shift = 1 Then
                     If InputBox("Ingrese la clave del Administrador" & Salto(1) & "Para ver toda la config. del sistema", "Seguridad", "") = "977611" Then
                        SSTab1.TabVisible(0) = True
                     End If
                  End If
                  
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaTodosLosDatosOP
' Description:       Procedimiento de cargar todos losvalores de los parametros
'
' Parameters :       Valor (Boolean)
'--------------------------------------------------------------------------------
Private Sub CargaTodosLosDatosOP(Valor As Boolean)
    Dim Pos As Integer
    On Error Resume Next
'    Pos = tdbgCuentas.Bookmark
    Call pLlenarParametros(Valor)
    Call pLlenarValoresParametros(Valor)
    
 '   tdbgCuentas.Bookmark = Pos

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
    
    pSel_Tab 0
    pSel_Tab2 0
    
    Call CargaTodosLosDatosOP(True)
    
    TabMantenimiento False
    TabMantenimiento2 False
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    SeteaBarraHerramientas Me.tbrOpcioneslibrocuenta, gsGrupo
    
    'tdbgParametros.MarqueeStyle = dbgHighlightCell
    tdbgValorParam.MarqueeStyle = dbgSolidCellBorder
    
    SSTab1.TabVisible(0) = True
    SSTab1.Tab = 1
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTab1
            .Width = Me.Width - .Left + 15 - 300
            .Height = Me.Height - .Top + 15 - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame4.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 400
        End With
       
        With tdbgCuentas
            .Width = Frame1.Width - .Left - 20
            .Height = Frame1.Height - .Top - 200
        End With
        
        With tdbgParametros
            .Width = Frame1.Width - .Left - 100
            .Height = Frame1.Height - .Top - 200
        End With
        
        sstParmetros.Width = Frame1.Width
        sstLibros.Width = Frame1.Width + 100
        
        Frame2.Width = sstParmetros.Width - 100
        Frame5.Width = sstLibros.Width - 500
        
        tdbgValorParam.Width = Frame2.Width - 100
        tdbgValorParam.Height = Frame2.Height - 500
        
        Frame3.Width = sstParmetros.Width - 800
        Frame7.Width = sstLibros.Width - 500
        
        tbrOpciones.Width = SSTab1.Width - 200
        tbrOpcioneslibrocuenta.Width = tbrOpciones.Width
        
        sstParmetros.Height = Me.Height - sstParmetros.Top - 700
        sstLibros.Height = sstParmetros.Height
        
        Frame2.Height = sstParmetros.Height - Frame2.Top - 200
        Frame5.Height = sstLibros.Height - Frame5.Top - 200
        
        tdbgLibro.Width = Frame5.Width - 300
        tdbgLibro.Height = Frame5.Height - 500
        
        Frame3.Height = Frame2.Height - 300
        Frame7.Height = Frame5.Height - 300
        
    End If
Exit Sub
errHand:
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       SSTab1_Click
' Description:       Evento que se ejecuta al hacer clic en el tab principal del formulario
'
' Parameters :       PreviousTab (Integer)
'--------------------------------------------------------------------------------
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        FiltrarRecordSetLibroCuenta
    Else
        FiltrarRecordSet
    End If
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       sstParmetros_KeyPress
' Description:       Evento que se ejecuta al hacer clic en el tab secundario del formulario
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub sstParmetros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sstParmetros.Tab = 1 Then
        pSetFocus tdbtValorParam
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tbrOpciones_ButtonClick
' Description:       Evento que se ejecuta al hacer clic en el toolbar de operaciones
'
' Parameters :       Button (MSComctlLib.Button)
'--------------------------------------------------------------------------------
Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo
        'Case 2: Ver
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
        Case 4: Borrar
        Case 5: Editar
        Case 6: ImprimirOp
        Case 7
            If Not sstParmetros.TabEnabled(1) Then
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar esta Operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
            End If
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ImprimirOp
' Description:       Procedimiento de impresion de configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ImprimirOp()
    Dim matriz(15) As Variant
    Dim Titulo As String
    Titulo = "Configuración de Operaciones"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;OPERACION;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;;True"
    matriz(5) = "@Titulo04;;True"
    matriz(6) = "@Titulo05;NUMERO DE CUENTA;True"
    matriz(7) = "@Titulo06;DESCRIPCION;True"
    matriz(8) = "@Titulo07;;True"
    matriz(9) = "@Tipo;CONFIG_OP;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"

    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandarAgrupado.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ImprimirLibCta
' Description:       Procedimiento de impresion de libors por cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ImprimirLibCta()
    Dim matriz(15) As Variant
    Dim Titulo As String
    Titulo = "Relacion Libro - Cuenta"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;DESCRIPCION DE LIBRO;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;;True"
    matriz(5) = "@Titulo04;;True"
    matriz(6) = "@Titulo05;NUMERO DE CUENTA;True"
    matriz(7) = "@Titulo06;DESCRIPCION;True"
    matriz(8) = "@Titulo07;;True"
    matriz(9) = "@Tipo;LIBRO_CTA;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"
    
    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"

    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandarAgrupado.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pSel_Tab
' Description:       Procedimiento de seleccion de tab principal
'
' Parameters :       nTab (Integer)
'--------------------------------------------------------------------------------
Sub pSel_Tab(nTab As Integer)
    Select Case nTab
        Case 0
            sstParmetros.TabEnabled(0) = True
            sstParmetros.TabEnabled(1) = False
            sstParmetros.Tab = 0
        Case 1
            sstParmetros.TabEnabled(0) = False
            sstParmetros.TabEnabled(1) = True
            sstParmetros.Tab = 1
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pSel_Tab2
' Description:       Procedimiento de seleccion de tab secundario
'
' Parameters :       nTab (Integer)
'--------------------------------------------------------------------------------
Sub pSel_Tab2(nTab As Integer)
    Select Case nTab
        Case 0
            sstLibros.TabEnabled(0) = True
            sstLibros.TabEnabled(1) = False
            sstLibros.Tab = 0
        Case 1
            sstLibros.TabEnabled(0) = False
            sstLibros.TabEnabled(1) = True
            sstLibros.Tab = 1
    End Select
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsValor)
    Call CerrarRecordSet(lrsParam)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ValidaNuevo
' Description:       Procedimiento que se ejecuta para validar un nuevo codigo de operacion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function ValidaNuevo() As Boolean
    ValidaNuevo = False
    
    If tdbgParametros.Columns(1).Text = "020" And CE(tdbgValorParam.Columns(2).Text) <> "" Then
        Mensajes "Solo se puede Ingresar/Modificar un tipo de cambio"
        pSetFocus tdbgValorParam
        Exit Function
    End If
    
    If tdbgParametros.Columns(1).Text = "027" And CE(tdbgValorParam.Columns(2).Text) <> "" Then
        Mensajes "Solo se puede Ingresar/Modificar una U.I.T."
        pSetFocus tdbgValorParam
        Exit Function
    End If
    
    ValidaNuevo = True
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ManNuevo
' Description:       Procedimiento para la creacion un nuevo codigo de operacion
'
' Parameters :
'--------------------------------------------------------------------------------
Sub ManNuevo()
    If CE(tdbgParametros) = "" Then
        Mensajes "Seleccione una operacion de la lista"
        Exit Sub
    End If
    
    If ValidaNuevo = False Then
        Exit Sub
    End If
    
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    'Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    pSel_Tab 1
    Call TabMantenimiento(True)
    tdbgParametros.Enabled = False
    pSetFocus tdbtValorParam
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ManNuevo2
' Description:       Procedimiento de creacion de un nuevo libro a la cuenta seleccionada
'
' Parameters :
'--------------------------------------------------------------------------------
Sub ManNuevo2()
    If CE(tdbgCuentas) = "" Then
        MsgBox "Seleccione una cuenta de la lista", vbOKOnly + vbInformation, gsNombreModulo
        Exit Sub
    End If

    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    'Call HabilitaControl(Me)
    ' ***
    lblMante2 = "NUEVO REGISTRO"
    pSel_Tab2 1
    Call TabMantenimiento2(True)
    tdbgCuentas.Enabled = False
    pSetFocus tdbtLibro
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Editar
' Description:       Procedimiento de edicion de configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Editar()
    If CE(tdbgParametros) = "" Then
        MsgBox "Seleccione una operacion de la lista", vbOKOnly + vbInformation, gsNombreModulo
        Exit Sub
    End If

    Call CargaDatosRegistro
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        'Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        pSel_Tab 1
        Call TabMantenimiento(True)
        tdbgParametros.Enabled = False
    Else
        lRegElim = False
    End If
    pSetFocus tdbtValorParam
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Editar2
' Description:       Procedimiento de edicion de los libros por cuneta
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Editar2()
    If CE(tdbgCuentas) = "" Then
        MsgBox "Seleccione una cuenta de la lista", vbOKOnly + vbInformation, gsNombreModulo
        Exit Sub
    End If

    Call CargaDatosRegistro2
    If lRegElim = False Then
       lTipoMnt = "EDITAR"
       If Me.lblMante2 = "VER REGISTRO" Then Call AseguraControl(Me, False)
       'Call HabilitaControl(Me)
       lblMante2 = "MODIFICANDO REGISTRO"
       pSel_Tab2 1
       Call TabMantenimiento2(True)
       tdbgCuentas.Enabled = False
    Else
       lRegElim = False
    End If
    pSetFocus tdbtLibro
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grabar
' Description:       Procedimiento de grabado de datos de las operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Grabar()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    
    If Not fValidarDatos Then Exit Sub
    
    If fValidaDuplicado2 Then Exit Sub
    
    Set clsMante = New clsMantoTablas
    ' *** Grabando Parametros
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt
    If lTipoMnt = "EDITAR" Then
        lArrMnt(0) = "ELIMINAR"
        lArrMnt(4) = tdbgValorParam.Columns(3).Text
    End If
    If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCND_CONFIG_OPERA", lArrMnt(), True) Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    If lTipoMnt = "EDITAR" Then
        lArrMnt(0) = "INSERTAR"
        lArrMnt(4) = tdbtValorParam.Text
        Set clsMante = New clsMantoTablas
        If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCND_CONFIG_OPERA", lArrMnt(), True) Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Exit Sub
            
        End If
    End If
    pLlenarValoresParametros
    FiltrarRecordSet
    Mensajes "Los datos se grabaron con exito...", vbInformation
    
    If (tdbgParametros.Columns(1).Value >= "012" And tdbgParametros.Columns(1).Value <= "023") Or _
       (tdbgParametros.Columns(1).Value >= "025" And tdbgParametros.Columns(1).Value <= "028") Or _
       (tdbgParametros.Columns(1).Value >= "029" And tdbgParametros.Columns(1).Value <= "046") Or _
       tdbgParametros.Columns(1).Value = "000" Then
       
            Call Cancelar
    Else
            If MsgBox("Desea Seguir ingresando más parámetros a este Concepto", vbYesNo + vbQuestion, gsNombreModulo) = vbYes Then
                tdbtValorParam = ""
                tdbtDescripción = ""
                pSetFocus tdbgValorParam
                pSetFocus tdbtValorParam
            Else
                Call Cancelar
            End If
    End If
        
    On Error Resume Next
    Dim nfila As Integer
    nfila = tdbgParametros.Bookmark
        
    Call CargaTodosLosDatosOP(True)
    Call FiltrarRecordSetLibroCuenta
    Call FiltrarRecordSet
    
    tdbgParametros.Bookmark = nfila
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Borrar
' Description:       Procedimiento de leiminar de las operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Borrar()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgValorParam.Columns(0).Value) <> "" Then
        respuesta = MsgBox("Desea Eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            lTipoMnt = "ELIMINAR"
            Call CargaArregloMnt
            lArrMnt(4) = tdbgValorParam.Columns(3).Value
            
            If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCND_CONFIG_OPERA", lArrMnt(), True) Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            'Call pLlenarValoresParametros
            Call CargaTodosLosDatosOP(True)
            FiltrarRecordSet
            Screen.MousePointer = vbDefault
            Mensajes "Registro ha sido eliminado", vbInformation
        End If
        pSetFocus tdbgValorParam
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Borrar2
' Description:       Procedimiento de eliminar los libros asociados a las cuentas
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Borrar2()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgLibro.Columns(0).Value) <> "" Then
       respuesta = MsgBox("Desea Eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
       If respuesta = vbYes Then
          Dim clsMante As clsMantoTablas
          Set clsMante = New clsMantoTablas
          ' *** Eliminando la Cuenta
          Screen.MousePointer = vbHourglass
          lTipoMnt = "ELIMINAR"
          Call CargaArregloMnt2
          lArrMnt2(3) = tdbgLibro.Columns(1).Value
          If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_RegLibroCta", lArrMnt2(), True) Then
             Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
             Screen.MousePointer = vbDefault
             Exit Sub
          End If
          Call pLlenarValoresParametros
          FiltrarRecordSetLibroCuenta
          Screen.MousePointer = vbDefault
          Mensajes "Registro ha sido eliminado", vbInformation
       End If
       pSetFocus tdbgLibro
    Else
       Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaDatosRegistro
' Description:       Procedimeinto de Carga de datos de las configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Sub CargaDatosRegistro()
    Me.tdbtValorParam = IIf(IsNull(tdbgValorParam.Columns(3).Value), "", tdbgValorParam.Columns(3).Value)
    Me.tdbtDescripción = IIf(IsNull(tdbgValorParam.Columns(4).Value), "", tdbgValorParam.Columns(4).Value)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaDatosRegistro2
' Description:       Procedimientos de carga los datos de los libros asociados a sus cuentas
'
' Parameters :
'--------------------------------------------------------------------------------
Sub CargaDatosRegistro2()
    tdbtLibro = IIf(IsNull(tdbgLibro.Columns(1).Value), "", tdbgLibro.Columns(1).Value)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fValidarDatos
' Description:       Funcion de validacion de configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Function fValidarDatos() As Boolean
    
    If CE(tdbgParametros) = "" Then
        MsgBox "Seleccione una operacion", vbOKOnly + vbInformation, gsNombreModulo
        fValidarDatos = False
        Exit Function
    End If
    
    If SSTab1.Tab = 1 Then
        If CE(tdbtLibro.Text) = "" Then
            MsgBox "Debe Ingresar un Valor de Parámetro para este concepto", vbOKOnly + vbInformation, gsNombreModulo
            fValidarDatos = False
            Exit Function
        End If
    
    Else
        If CE(tdbtValorParam.Text) = "" Then
            MsgBox "Debe Ingresar un Valor de Parámetro para este concepto", vbOKOnly + vbInformation, gsNombreModulo
            fValidarDatos = False
            Exit Function
        End If
    End If
    
    
    fValidarDatos = True
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fValidarDatos2
' Description:       Funcion de validacion de cuentas por libro
'
' Parameters :
'--------------------------------------------------------------------------------
Function fValidarDatos2() As Boolean

    If CE(tdbgCuentas) = "" Then
        MsgBox "Seleccione una cuenta de la lista", vbOKOnly + vbInformation, gsNombreModulo
        fValidarDatos2 = False
        Exit Function
    End If

    If CE(tdbtLibro.Text) = "" Then
        MsgBox "Debe Ingresar el Libro", vbOKOnly + vbInformation, gsNombreModulo
        fValidarDatos2 = False
        Exit Function
    End If
    fValidarDatos2 = True
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fValidaDuplicado
' Description:       Funcion de validacion de duplicado de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Function fValidaDuplicado() As Boolean
    Dim objCon As clsMantoTablas
    Dim rs As ADODB.Recordset
    Dim arr() As Variant
    Dim sql As String
    sql = "spCND_CONFIG_OPERA 'BUSCARREGISTRO','" & gsEmpresa & "','" & gsAnio & "','" & _
           tdbgParametros.Columns(1).Value & "','" & tdbtValorParam.Text & "'"
    arr = Array(sql)
    Set objCon = New clsMantoTablas
    Set rs = objCon.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
    If rs Is Nothing Then
        fValidaDuplicado = False: Exit Function
    End If
    Mensajes "Código ya Existe"
    fValidaDuplicado = True
    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fValidaDuplicado2
' Description:       Funcion de duplicado de cuentas por libro
'
' Parameters :
'--------------------------------------------------------------------------------
Function fValidaDuplicado2() As Boolean
    Dim objCon As clsMantoTablas
    Dim rs As ADODB.Recordset
    Dim arr() As Variant
    Dim sql As String
    sql = "spCn_RegLibroCta 'BUSCARREGISTRO','" & gsEmpresa & "','" & tdbgCuentas.Columns(0).Value & "','" & tdbtLibro.Text & "','" & gsAnio & "'"
    arr = Array(sql)
    Set objCon = New clsMantoTablas
    Set rs = objCon.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
    If rs Is Nothing Then
        fValidaDuplicado2 = False: Exit Function
    End If
    Mensajes "Código ya Existe"
    fValidaDuplicado2 = True
    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       FiltrarRecordSet
' Description:       Filtrado de registros de la configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(0) As String
    Dim i As Integer
    If lrsValor Is Nothing Then Exit Sub
    cadena = ""
    If Not IsNull(tdbgParametros.Columns(1).Value) Then filtros(0) = "Cop_cCodigo like '" & tdbgParametros.Columns(1).Value & "'"
    For i = 0 To 0
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    If Trim(cadena) <> "" Then
        lrsValor.Filter = cadena
    Else
        lrsValor.Filter = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       TabMantenimiento
' Description:       Procedimiento de activacion de botones del toolbar de operaciones
'
' Parameters :       Valor (Boolean)
'--------------------------------------------------------------------------------
Private Sub TabMantenimiento(Valor As Boolean)
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    'tbrOpciones.Buttons(2).Enabled = Not valor  ' *** Buscar
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
    If Valor = True Then
        tbrOpciones.Buttons(7).Image = 8
    Else
        tbrOpciones.Buttons(7).Image = 7
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       TabMantenimiento2
' Description:       Procedimiento de activacion de botones del toolbar de cuentas por libro
'
' Parameters :       Valor (Boolean)
'--------------------------------------------------------------------------------
Private Sub TabMantenimiento2(Valor As Boolean)
    tbrOpcioneslibrocuenta.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    'tbrOpciones.Buttons(2).Enabled = Not valor  ' *** Buscar
    tbrOpcioneslibrocuenta.Buttons(3).Enabled = Valor      ' *** Grabar
    tbrOpcioneslibrocuenta.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpcioneslibrocuenta.Buttons(5).Enabled = Not Valor  ' *** Editar
    If Valor = True Then
        tbrOpcioneslibrocuenta.Buttons(7).Image = 8
    Else
        tbrOpcioneslibrocuenta.Buttons(7).Image = 7
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Cancelar
' Description:       Procedimiento de cancelacion de configuracion de operaciones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Cancelar()
    tdbtValorParam = ""
    tdbtDescripción = ""
    'If Me.lblMante = "VER REGISTRO" Then
    '    Call AseguraControl(Me, False)
    'Else
    '    Call HabilitaControl(Me)
    'End If
    pSel_Tab 0
    Call TabMantenimiento(False)
    tdbgParametros.Enabled = True
    pSetFocus tdbgValorParam
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Cancelar2
' Description:       Procedimiento de cancelacion de cuentas por libro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Cancelar2()
    tdbtLibro = ""
    tdbtDescripcion = ""
    'If Me.lblMante = "VER REGISTRO" Then
    '    Call AseguraControl(Me, False)
    'Else
    '    Call HabilitaControl(Me)
    'End If
    pSel_Tab2 0
    Call TabMantenimiento2(False)
    tdbgCuentas.Enabled = True
    pSetFocus tdbgLibro
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaArregloMnt
' Description:       Procedimiento de llenar el arreglo de configuracion de oepraciones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CargaArregloMnt()
    ReDim lArrMnt(8) As Variant
    lArrMnt(0) = lTipoMnt                           ' Accion
    lArrMnt(1) = gsEmpresa                          ' Empresa
    lArrMnt(2) = gsAnio                             ' Año
    lArrMnt(3) = tdbgParametros.Columns(1).Value    ' Codigo
    lArrMnt(4) = tdbtValorParam.Text                ' Valor Parametro
    lArrMnt(5) = tdbnIGV.Value                      ' Valor Igv
    lArrMnt(6) = "A"                                ' Estado
    lArrMnt(7) = gsUsuario                          ' Usuario
    lArrMnt(8) = gsUsuario                          ' Usuario
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaArregloMnt2
' Description:       Procedimiento de llenar el arreglo de cuentas por libro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CargaArregloMnt2()
    ReDim lArrMnt2(4) As Variant
    lArrMnt2(0) = lTipoMnt                           ' Accion
    lArrMnt2(1) = gsEmpresa                          ' Empresa
    lArrMnt2(2) = tdbgCuentas.Columns(0).Value       ' Cuenta
    lArrMnt2(3) = tdbgLibro.Columns(1).Text          ' Libro
    lArrMnt2(4) = gsAnio                             ' Año
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tbrOpcioneslibrocuenta_ButtonClick
' Description:       Evento que se ejecuta al hacer clic en los botones del toolbar de libros por cuenta
'
' Parameters :       Button (MSComctlLib.Button)
'--------------------------------------------------------------------------------
Private Sub tbrOpcioneslibrocuenta_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo2
        'Case 2: Ver
        Case 3: Grabar2
                SeteaBarraHerramientas Me.tbrOpcioneslibrocuenta, gsGrupo
        Case 4: Borrar2
        Case 5: Editar2
        Case 6: ImprimirLibCta
        Case 7
            If Not sstLibros.TabEnabled(1) Then
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar esta Operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar2
                    SeteaBarraHerramientas Me.tbrOpcioneslibrocuenta, gsGrupo
                End If
            End If
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgCuentas_HeadClick
' Description:       Evento que se ejecuta al hacer clic en la cabecera de la grilla
'
' Parameters :       ColIndex (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgCuentas_HeadClick(ByVal ColIndex As Integer)
If Not lrsParam2 Is Nothing Then
    If lrsParam2.RecordCount > 0 Then
    
        lrsParam2.Sort = tdbgCuentas.Columns(ColIndex).DataField
        tdbgCuentas.DataSource = lrsParam2
        
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgCuentas_RowColChange
' Description:       Evento que se ejecuta al cambiar el cursor en las filas de la grilla
'
' Parameters :       LastRow (Variant)
'                    LastCol (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgCuentas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    FiltrarRecordSetLibroCuenta
End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgLibro_HeadClick
' Description:       Evento que se ejecuta al hacer clic enla cabecera de la grilla de libros
'
' Parameters :       ColIndex (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgLibro_HeadClick(ByVal ColIndex As Integer)
If Not lrsLibro Is Nothing Then
    If lrsLibro.RecordCount > 0 Then
    
        lrsLibro.Sort = tdbgLibro.Columns(ColIndex).DataField
        tdbgLibro.DataSource = lrsLibro
        
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgParametros_HeadClick
' Description:       Evento que se ejecuta al hacer clic en la cabecera de la grilla de parametros
'
' Parameters :       ColIndex (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgParametros_HeadClick(ByVal ColIndex As Integer)
If Not lrsParam Is Nothing Then
    If lrsParam.RecordCount > 0 Then
    
        lrsParam.Sort = tdbgParametros.Columns(ColIndex).DataField
        tdbgParametros.DataSource = lrsParam
        
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgParametros_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla la grulla de parametros
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgParametros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sstParmetros.Tab = 0 Then
        pSetFocus tdbgValorParam
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgParametros_RowColChange
' Description:       Evento que se ejecuta al cambiar de fila en la grilla de parametros
'
' Parameters :       LastRow (Variant)
'                    LastCol (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgParametros_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If tdbgParametros.Columns(1).Value <> "004" And tdbgParametros.Columns(1).Value <> "005" And _
        tdbgParametros.Columns(1).Value <> "006" And tdbgParametros.Columns(1).Value <> "007" Then
        tdbgValorParam.Columns(4).Width = 4845
        tdbgValorParam.Columns(5).AllowFocus = False
        tdbgValorParam.Columns(5).Visible = False
        lblIgv.Visible = False
        tdbnIGV.Visible = False
    Else
        tdbgValorParam.Columns(4).Width = 4310
        tdbgValorParam.Columns(5).AllowFocus = False
        tdbgValorParam.Columns(5).Visible = False
        'lblIgv.Visible = True
        'tdbnIGV.Visible = True
    End If
    If tdbgParametros.Columns(3).Value = "CTA" Or tdbgParametros.Columns(3).Value = "VAL" Then
        tdbtValorParam.MaxLength = 12
    ElseIf tdbgParametros.Columns(3).Value = "TPC" Then
        
       tdbtValorParam.MaxLength = 3
    Else
        tdbtValorParam.MaxLength = 2
    End If
    
    If tdbgParametros.Columns(3).Value = "VAL" Then
        tdbgValorParam.Columns(4).Visible = False 'DESCRIPCION
        tdbtDescripción.Visible = False
    Else
        tdbgValorParam.Columns(4).Visible = True 'DESCRIPCION
        tdbtDescripción.Visible = True
    End If
    
    FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgValorParam_HeadClick
' Description:       Evento que se ejecuta al hacer clic en la cabecera de la grilla devalores de parametros
'
' Parameters :       ColIndex (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgValorParam_HeadClick(ByVal ColIndex As Integer)
If Not lrsValor Is Nothing Then
    If lrsValor.RecordCount > 0 Then
    
        lrsValor.Sort = tdbgValorParam.Columns(ColIndex).DataField
        tdbgValorParam.DataSource = lrsValor
        
    End If
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgValorParam_KeyPress
' Description:       Evento que se ejecuta al presioanr una tecla en la grilla de valores de parametros
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgValorParam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Editar
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtDescripcion_GotFocus
' Description:       Evento que se ejecuta al salor del enfoque del campo de descripcion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtDescripcion_GotFocus()
    If tdbgParametros.Columns(3).Value = "DOC" Or tdbgParametros.Columns(3).Value = "CTA" Then
       tdbtDescripcion.ReadOnly = True
    Else
        tdbtDescripcion.ReadOnly = False
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtDescripción_GotFocus
' Description:       Evento que se ejecuta al recibir el enfoque en el campo de descripcion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtDescripción_GotFocus()
    If tdbgParametros.Columns(3).Value = "DOC" Or tdbgParametros.Columns(3).Value = "CTA" Then
       tdbtDescripción.ReadOnly = True
    Else
        tdbtDescripción.ReadOnly = False
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibro_Change
' Description:       Evento que se ejecuta al cambiar el libro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtLibro_Change()
   pTextChange tdbtLibro, tdbtDescripcion
   pSetFocus tdbtLibro
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibro_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el libro
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "Libro", Control, "Libros", Me, gsPeriodo, LTrim(tdbtLibro.Text))
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtValorParam_Change
' Description:       Evento que se ejecuta al cambiar el texto de los valores de parametros
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtValorParam_Change()
    If tdbgParametros.Columns(3).Value = "DOC" Or tdbgParametros.Columns(3).Value = "TPC" Then
        If Len(tdbtValorParam) < tdbtValorParam.MaxLength Then
            tdbtDescripción.Text = ""
        Else
            tdbtValorParam_LostFocus
        End If
    ElseIf tdbgParametros.Columns(3).Value = "VAL" Then
        tdbtDescripción.Text = ""
    Else
        If Len(tdbtValorParam) <= 1 Then
            tdbtDescripción.Text = ""
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtValorParam_KeyPress
' Description:       Evento que se ejecuta al presioanr una tecla en los valores de parametros
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtValorParam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(tdbtValorParam) > 1 Then
            tdbtValorParam_LostFocus
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtValorParam_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque en los valores de parametros
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtValorParam_LostFocus()
    If tdbgParametros.Columns(3).Value = "CTA" Then
        If Len(tdbtValorParam) <= 1 Then Exit Sub
        sqlSp = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as Descripcion, pla_cTitulo From dbo.CNM_PLAN_CTA " & _
                "Where Emp_cCodigo='" & gsEmpresa & "' AND Pan_cAnio='" & gsAnio & "' AND Pla_cCuentaContable='" & tdbtValorParam.Text & _
                "' AND Pla_cEstado='A' ORDER BY Pla_cCuentaContable "
        tdbtDescripción = fDevuelveDescripcion(sqlSp)
        If tdbtDescripción.Text = "" Then tdbtValorParam.Text = ""
        pSetFocus tdbtValorParam
        
    ElseIf tdbgParametros.Columns(3).Value = "TPC" Then
             If Trim(tdbtValorParam.Text) <> "" Then
                sqlSp = "SELECT     Tab_cCodigo AS Codigo, Tab_cDescripCampo AS Descripcion" _
                     & " From tabla " _
                    & " WHERE     (Emp_cCodigo = '" & gsEmpresa & "') AND (Tab_cTabla = '026') AND  Tab_cCodigo = '" & tdbtValorParam & "'"
                tdbtDescripción = fDevuelveDescripcion(sqlSp)
                If tdbtDescripción.Text = "" Then tdbtValorParam.Text = ""
                pSetFocus tdbtValorParam
            End If
    Else
    
        If CE(tdbtValorParam.Text) <> "" Then
            If tdbgParametros.Columns(3).Value <> "VAL" Then
                sqlSp = "spCn_ConsultaTipDocsLibro 'SEL_DOCS_REG','" & gsEmpresa & "','','','" & tdbtValorParam.Text & "'"
                tdbtDescripción.Text = fDevuelveDescripcion(sqlSp)
                If tdbtDescripción.Text = "" Then tdbtValorParam.Text = ""
                pSetFocus tdbtValorParam
            End If
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtValorParam_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en los valores de parametros
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtValorParam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        tdbtValorParam_LostFocus
        DoEvents
        
        If tdbgParametros.Columns(3).Value = "CTA" Then
            Call LlamaBuscar(frmBuscador, "CTA", Control, "Cuentas", Me, gsPeriodo, CE(tdbtValorParam.Text))
        ElseIf tdbgParametros.Columns(3).Value = "DOC" Then
            Call LlamaBuscar(frmBuscador, "DOC", Control, "TipoDocumento", Me, gsPeriodo)
        ElseIf tdbgParametros.Columns(3).Value = "TPC" Then
            Call LlamaBuscar(frmBuscador, "TPC", Control, "TipoCambio", Me, gsPeriodo)
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       RecibirDatos
' Description:       Procedimiento que s eutiliza para recibir los datos del forumario de busqueda
'
' Parameters :       lControl (String)
'                    param0 (String)
'                    param1 (String)
'                    param2 (String)
'--------------------------------------------------------------------------------
Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
    Select Case Control
           Case "CTA"
                tdbtValorParam = Trim(param0)
                tdbtDescripción = Trim(param1)
                pSetFocus tdbtValorParam
           Case "Libro"
                tdbtLibro = Trim(param0)
                tdbtDescripcion = Trim(param1)
                pSetFocus tdbtLibro

           Case "DOC"
                tdbtValorParam = Trim(param0)
                tdbtDescripción = Trim(param1)
                pSetFocus tdbtValorParam
            
    End Select
    Unload frmBuscador
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fDevuelveDescripcion
' Description:       Funcion que retorna la descripcion de la consulta ingresada como parametro
'
' Parameters :       Sql (String)
'--------------------------------------------------------------------------------
Function fDevuelveDescripcion(sql As String) As String
Dim objCon As clsMantoTablas
Dim rs As ADODB.Recordset
Dim arr() As Variant
arr = Array(sql)
Set objCon = New clsMantoTablas
Set rs = objCon.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
If rs Is Nothing Then
    Mensajes "Código No Existe"
    fDevuelveDescripcion = ""
    Exit Function
End If
fDevuelveDescripcion = NuloText(rs!Descripcion)
rs.Close
Set rs = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       FiltrarRecordSetLibroCuenta
' Description:       Procedimiento de filtrado de libro por cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub FiltrarRecordSetLibroCuenta()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(0) As String
    Dim i As Integer
    If lrsLibro Is Nothing Then Exit Sub
    cadena = ""
    If Not IsNull(tdbgCuentas.Columns(0).Value) Then filtros(0) = "Pla_cCuentaContable like '" & tdbgCuentas.Columns(0).Value & "'"
    For i = 0 To 0
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    If Trim(cadena) <> "" Then
        lrsLibro.Filter = cadena
    Else
        lrsLibro.Filter = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grabar2
' Description:       Procedimiento de grabar las configuracion de libro por cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Sub Grabar2()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    
    If Not fValidarDatos2 Then Exit Sub
    
    If fValidaDuplicado2 Then Exit Sub
    
    Set clsMante = New clsMantoTablas
    ' *** Grabando Parametros
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt2
    
    If lTipoMnt = "EDITAR" Then
       lArrMnt2(0) = "ELIMINAR"
       lArrMnt2(3) = tdbgLibro.Columns(1).Text
        If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_RegLibroCta", lArrMnt2(), True) Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Exit Sub
        End If
    End If
    
    If lTipoMnt = "INSERTAR" Or lTipoMnt = "EDITAR" Then
        lArrMnt2(0) = "INSERTAR"
        lArrMnt2(3) = tdbtLibro.Text
        Set clsMante = New clsMantoTablas
        If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_RegLibroCta", lArrMnt2(), True) Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Dim nfila As Integer
    nfila = tdbgParametros.Bookmark
        
    Call pLlenarValoresParametros
    Call FiltrarRecordSetLibroCuenta
    Call FiltrarRecordSet
    
    tdbgParametros.Bookmark = nfila
    
    
    Mensajes "Los datos se grabaron con exito...", vbInformation
    Call Cancelar2
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pTextChange
' Description:       Procedimiento que se ejecuta al cambiar el texto de un TextBox
'
' Parameters :       oTxtCodigo (TDBText)
'                    oTxtDescrip (TDBText)
'--------------------------------------------------------------------------------
Private Sub pTextChange(oTxtCodigo As TDBText, oTxtDescrip As TDBText)
    If Len(oTxtCodigo) < oTxtCodigo.MaxLength Then
        oTxtDescrip.Text = ""
    Else
        pTextLostFocus oTxtCodigo, oTxtDescrip
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pTextLostFocus
' Description:       Procedimiento que se ejecuta al salir del enfoque de un textbox
'
' Parameters :       oTxtCodigo (TDBText)
'                    oTxtDescrip (TDBText)
'--------------------------------------------------------------------------------
Private Sub pTextLostFocus(oTxtCodigo As TDBText, oTxtDescrip As TDBText)
    If Len(oTxtCodigo) < 2 Then Exit Sub
    sqlSp = "spCn_GrabaLibroOpera 'BUSCARREGISTRO','" & gsEmpresa & "','" & gsAnio & "','" & oTxtCodigo.Text & "'"
    oTxtDescrip = fDevuelveDescripcion(sqlSp)
End Sub

