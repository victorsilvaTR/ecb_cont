VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta de Valores"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13035
   Icon            =   "frmManValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   13035
   Begin VB.Frame fraLeyenda 
      Height          =   600
      Left            =   225
      TabIndex        =   13
      Top             =   1035
      Width           =   12690
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Voucher"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   6705
         TabIndex        =   19
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor unitario y cantidad invalida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   7605
         TabIndex        =   18
         Top             =   270
         Width           =   2820
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Voucher con Importe modificado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3690
         TabIndex        =   17
         Top             =   270
         Width           =   2805
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Voucher"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2790
         TabIndex        =   16
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Eliminados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   15
         Top             =   270
         Width           =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Voucher"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   225
         Width           =   675
      End
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   1170
      TabIndex        =   0
      Top             =   225
      Width           =   3165
      _ExtentX        =   5583
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
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
      _PropDict       =   $"frmManValores.frx":0ECA
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
   Begin TrueOleDBGrid70.TDBGrid grdCapital 
      Height          =   4665
      Left            =   195
      TabIndex        =   8
      Top             =   1710
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8229
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Voucher"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Item"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Entidad"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Codigo"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Apellidos y Nombres, Razón  Social"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "codigo"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Titulo - Denominación"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6).DropDown=   "tdbdTipoAcciones"
      Columns(6).DropDown.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Valor nominal unitario"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "External Editor"
      Columns(7).ExternalEditor=   "TDBNumberNeg"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Cantidad"
      Columns(8).DataField=   ""
      Columns(8).DefaultValue=   "0.00"
      Columns(8).DefaultValue.vt=   8
      Columns(8).NumberFormat=   "External Editor"
      Columns(8).ExternalEditor=   "TDBNumberNeg"
      Columns(8).ExternalEditor.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Costo Total"
      Columns(9).DataField=   ""
      Columns(9).DefaultValue=   "0"
      Columns(9).DefaultValue.vt=   8
      Columns(9).NumberFormat=   "External Editor"
      Columns(9).ExternalEditor=   "TDBNumberNeg"
      Columns(9).ExternalEditor.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Provision Total"
      Columns(10).DataField=   ""
      Columns(10).DefaultValue=   "0"
      Columns(10).DefaultValue.vt=   8
      Columns(10).NumberFormat=   "External Editor"
      Columns(10).ExternalEditor=   "TDBNumberNeg"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Otros Costos"
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "External Editor"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Ajuste al Valor Razonable"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Total Neto"
      Columns(13).DataField=   ""
      Columns(13).DefaultValue=   "0"
      Columns(13).DefaultValue.vt=   8
      Columns(13).NumberFormat=   "External Editor"
      Columns(13).ExternalEditor=   "TDBNumberNeg"
      Columns(13).ExternalEditor.vt=   8
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "FLAG"
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=1085"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1005"
      Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=139777"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=4180"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=4101"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=139780"
      Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=131588"
      Splits(0)._ColumnProps(41)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(44)=   "Column(6).Width=2487"
      Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=2408"
      Splits(0)._ColumnProps(47)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=131586"
      Splits(0)._ColumnProps(49)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(51)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(52)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._ColStyle=131588"
      Splits(0)._ColumnProps(57)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(58)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(59)=   "Column(8).Width=2699"
      Splits(0)._ColumnProps(60)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(8)._WidthInPix=2619"
      Splits(0)._ColumnProps(62)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(63)=   "Column(8)._ColStyle=197122"
      Splits(0)._ColumnProps(64)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(65)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(66)=   "Column(9).Width=2778"
      Splits(0)._ColumnProps(67)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(9)._WidthInPix=2699"
      Splits(0)._ColumnProps(69)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(70)=   "Column(9)._ColStyle=197122"
      Splits(0)._ColumnProps(71)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(72)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(73)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(74)=   "Column(10).Width=2302"
      Splits(0)._ColumnProps(75)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(10)._WidthInPix=2223"
      Splits(0)._ColumnProps(77)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(78)=   "Column(10)._ColStyle=197122"
      Splits(0)._ColumnProps(79)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(80)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(81)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(82)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(83)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=131588"
      Splits(0)._ColumnProps(86)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(87)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(88)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(89)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(90)=   "Column(12)._ColStyle=131588"
      Splits(0)._ColumnProps(91)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(92)=   "Column(13).Width=2514"
      Splits(0)._ColumnProps(93)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(13)._WidthInPix=2434"
      Splits(0)._ColumnProps(95)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(96)=   "Column(13)._ColStyle=197122"
      Splits(0)._ColumnProps(97)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(98)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(99)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(100)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(101)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(103)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(104)=   "Column(14)._ColStyle=131588"
      Splits(0)._ColumnProps(105)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(107)=   "Column(14).Order=15"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=15"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=197124"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=197124"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=197124"
      Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=1085"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=1005"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=205313"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=9128"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=9049"
      Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=205316"
      Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(41)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(42)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(44)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(45)=   "Column(5)._ColStyle=197120"
      Splits(1)._ColumnProps(46)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(47)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(48)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(49)=   "Column(6).Width=2646"
      Splits(1)._ColumnProps(50)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(51)=   "Column(6)._WidthInPix=2566"
      Splits(1)._ColumnProps(52)=   "Column(6)._ColStyle=197120"
      Splits(1)._ColumnProps(53)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(54)=   "Column(6).AutoDropDown=1"
      Splits(1)._ColumnProps(55)=   "Column(6).DropDownList=1"
      Splits(1)._ColumnProps(56)=   "Column(7).Width=2090"
      Splits(1)._ColumnProps(57)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(58)=   "Column(7)._WidthInPix=2011"
      Splits(1)._ColumnProps(59)=   "Column(7)._ColStyle=197122"
      Splits(1)._ColumnProps(60)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(61)=   "Column(8).Width=1429"
      Splits(1)._ColumnProps(62)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(63)=   "Column(8)._WidthInPix=1349"
      Splits(1)._ColumnProps(64)=   "Column(8)._ColStyle=197122"
      Splits(1)._ColumnProps(65)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(66)=   "Column(9).Width=1931"
      Splits(1)._ColumnProps(67)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(68)=   "Column(9)._WidthInPix=1852"
      Splits(1)._ColumnProps(69)=   "Column(9)._ColStyle=197122"
      Splits(1)._ColumnProps(70)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(71)=   "Column(10).Width=1958"
      Splits(1)._ColumnProps(72)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(10)._WidthInPix=1879"
      Splits(1)._ColumnProps(74)=   "Column(10)._ColStyle=197122"
      Splits(1)._ColumnProps(75)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(76)=   "Column(11).Width=1508"
      Splits(1)._ColumnProps(77)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(78)=   "Column(11)._WidthInPix=1429"
      Splits(1)._ColumnProps(79)=   "Column(11)._ColStyle=197124"
      Splits(1)._ColumnProps(80)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(81)=   "Column(12).Width=1799"
      Splits(1)._ColumnProps(82)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(12)._WidthInPix=1720"
      Splits(1)._ColumnProps(84)=   "Column(12)._ColStyle=197124"
      Splits(1)._ColumnProps(85)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(86)=   "Column(13).Width=767"
      Splits(1)._ColumnProps(87)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(88)=   "Column(13)._WidthInPix=688"
      Splits(1)._ColumnProps(89)=   "Column(13)._ColStyle=197122"
      Splits(1)._ColumnProps(90)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(91)=   "Column(14).Width=1667"
      Splits(1)._ColumnProps(92)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(14)._WidthInPix=1588"
      Splits(1)._ColumnProps(94)=   "Column(14).AllowSizing=0"
      Splits(1)._ColumnProps(95)=   "Column(14)._ColStyle=197124"
      Splits(1)._ColumnProps(96)=   "Column(14).Visible=0"
      Splits(1)._ColumnProps(97)=   "Column(14).AllowFocus=0"
      Splits(1)._ColumnProps(98)=   "Column(14).Order=15"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   3
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=37,.bgcolor=&HFFDBBB&,.bold=0"
      _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bgcolor=&HD2D2D2&"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=130,.parent=13,.alignment=2,.bgcolor=&HF8ECC9&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=127,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=128,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=129,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=138,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=135,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=136,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=137,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=122,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=119,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=120,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=121,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=98,.parent=13,.alignment=2,.bgcolor=&HF8ECC9&"
      _StyleDefs(50)  =   ":id=98,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=95,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=96,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=97,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.bgcolor=&HF8ECC9&,.wraptext=-1"
      _StyleDefs(55)  =   ":id=32,.locked=-1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=146,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=143,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=144,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=145,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15,.alignment=1"
      _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(9).Style:id=62,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15,.alignment=1"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15,.alignment=1"
      _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=162,.parent=13"
      _StyleDefs(84)  =   "Splits(0).Columns(11).HeadingStyle:id=159,.parent=14"
      _StyleDefs(85)  =   "Splits(0).Columns(11).FooterStyle:id=160,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(11).EditorStyle:id=161,.parent=17"
      _StyleDefs(87)  =   "Splits(0).Columns(12).Style:id=54,.parent=13"
      _StyleDefs(88)  =   "Splits(0).Columns(12).HeadingStyle:id=51,.parent=14"
      _StyleDefs(89)  =   "Splits(0).Columns(12).FooterStyle:id=52,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(12).EditorStyle:id=53,.parent=17"
      _StyleDefs(91)  =   "Splits(0).Columns(13).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(92)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
      _StyleDefs(93)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15,.alignment=1"
      _StyleDefs(94)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(95)  =   "Splits(0).Columns(14).Style:id=154,.parent=13"
      _StyleDefs(96)  =   "Splits(0).Columns(14).HeadingStyle:id=151,.parent=14"
      _StyleDefs(97)  =   "Splits(0).Columns(14).FooterStyle:id=152,.parent=15"
      _StyleDefs(98)  =   "Splits(0).Columns(14).EditorStyle:id=153,.parent=17"
      _StyleDefs(99)  =   "Splits(1).Style:id=25,.parent=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(100) =   "Splits(1).CaptionStyle:id=76,.parent=4"
      _StyleDefs(101) =   "Splits(1).HeadingStyle:id=26,.parent=2"
      _StyleDefs(102) =   "Splits(1).FooterStyle:id=27,.parent=3,.alignment=1"
      _StyleDefs(103) =   "Splits(1).InactiveStyle:id=28,.parent=5"
      _StyleDefs(104) =   "Splits(1).SelectedStyle:id=44,.parent=6"
      _StyleDefs(105) =   "Splits(1).EditorStyle:id=43,.parent=7"
      _StyleDefs(106) =   "Splits(1).HighlightRowStyle:id=45,.parent=8"
      _StyleDefs(107) =   "Splits(1).EvenRowStyle:id=46,.parent=9"
      _StyleDefs(108) =   "Splits(1).OddRowStyle:id=75,.parent=10"
      _StyleDefs(109) =   "Splits(1).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(110) =   "Splits(1).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(111) =   "Splits(1).Columns(0).Style:id=134,.parent=25"
      _StyleDefs(112) =   "Splits(1).Columns(0).HeadingStyle:id=131,.parent=26"
      _StyleDefs(113) =   "Splits(1).Columns(0).FooterStyle:id=132,.parent=27"
      _StyleDefs(114) =   "Splits(1).Columns(0).EditorStyle:id=133,.parent=43"
      _StyleDefs(115) =   "Splits(1).Columns(1).Style:id=142,.parent=25"
      _StyleDefs(116) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=26"
      _StyleDefs(117) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=27"
      _StyleDefs(118) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=43"
      _StyleDefs(119) =   "Splits(1).Columns(2).Style:id=126,.parent=25"
      _StyleDefs(120) =   "Splits(1).Columns(2).HeadingStyle:id=123,.parent=26"
      _StyleDefs(121) =   "Splits(1).Columns(2).FooterStyle:id=124,.parent=27"
      _StyleDefs(122) =   "Splits(1).Columns(2).EditorStyle:id=125,.parent=43"
      _StyleDefs(123) =   "Splits(1).Columns(3).Style:id=82,.parent=25,.alignment=2,.bgcolor=&HF8ECC9&"
      _StyleDefs(124) =   ":id=82,.locked=-1"
      _StyleDefs(125) =   "Splits(1).Columns(3).HeadingStyle:id=79,.parent=26"
      _StyleDefs(126) =   "Splits(1).Columns(3).FooterStyle:id=80,.parent=27"
      _StyleDefs(127) =   "Splits(1).Columns(3).EditorStyle:id=81,.parent=43"
      _StyleDefs(128) =   "Splits(1).Columns(4).Style:id=86,.parent=25,.bgcolor=&HF8ECC9&,.locked=-1"
      _StyleDefs(129) =   "Splits(1).Columns(4).HeadingStyle:id=83,.parent=26"
      _StyleDefs(130) =   "Splits(1).Columns(4).FooterStyle:id=84,.parent=27"
      _StyleDefs(131) =   "Splits(1).Columns(4).EditorStyle:id=85,.parent=43"
      _StyleDefs(132) =   "Splits(1).Columns(5).Style:id=150,.parent=25,.alignment=0"
      _StyleDefs(133) =   "Splits(1).Columns(5).HeadingStyle:id=147,.parent=26"
      _StyleDefs(134) =   "Splits(1).Columns(5).FooterStyle:id=148,.parent=27"
      _StyleDefs(135) =   "Splits(1).Columns(5).EditorStyle:id=149,.parent=43"
      _StyleDefs(136) =   "Splits(1).Columns(6).Style:id=90,.parent=25,.alignment=0,.bgcolor=&HFFFFFF&"
      _StyleDefs(137) =   "Splits(1).Columns(6).HeadingStyle:id=87,.parent=26"
      _StyleDefs(138) =   "Splits(1).Columns(6).FooterStyle:id=88,.parent=27,.bgcolor=&HD2D2D2&"
      _StyleDefs(139) =   "Splits(1).Columns(6).EditorStyle:id=89,.parent=43"
      _StyleDefs(140) =   "Splits(1).Columns(7).Style:id=118,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(141) =   "Splits(1).Columns(7).HeadingStyle:id=115,.parent=26"
      _StyleDefs(142) =   "Splits(1).Columns(7).FooterStyle:id=116,.parent=27"
      _StyleDefs(143) =   "Splits(1).Columns(7).EditorStyle:id=117,.parent=43"
      _StyleDefs(144) =   "Splits(1).Columns(8).Style:id=102,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(145) =   "Splits(1).Columns(8).HeadingStyle:id=99,.parent=26"
      _StyleDefs(146) =   "Splits(1).Columns(8).FooterStyle:id=100,.parent=27"
      _StyleDefs(147) =   "Splits(1).Columns(8).EditorStyle:id=101,.parent=43,.alignment=1"
      _StyleDefs(148) =   "Splits(1).Columns(9).Style:id=106,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(149) =   ":id=106,.locked=0"
      _StyleDefs(150) =   "Splits(1).Columns(9).HeadingStyle:id=103,.parent=26"
      _StyleDefs(151) =   "Splits(1).Columns(9).FooterStyle:id=104,.parent=27,.alignment=1"
      _StyleDefs(152) =   "Splits(1).Columns(9).EditorStyle:id=105,.parent=43"
      _StyleDefs(153) =   "Splits(1).Columns(10).Style:id=110,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(154) =   ":id=110,.locked=0"
      _StyleDefs(155) =   "Splits(1).Columns(10).HeadingStyle:id=107,.parent=26"
      _StyleDefs(156) =   "Splits(1).Columns(10).FooterStyle:id=108,.parent=27"
      _StyleDefs(157) =   "Splits(1).Columns(10).EditorStyle:id=109,.parent=43"
      _StyleDefs(158) =   "Splits(1).Columns(11).Style:id=166,.parent=25"
      _StyleDefs(159) =   "Splits(1).Columns(11).HeadingStyle:id=163,.parent=26"
      _StyleDefs(160) =   "Splits(1).Columns(11).FooterStyle:id=164,.parent=27"
      _StyleDefs(161) =   "Splits(1).Columns(11).EditorStyle:id=165,.parent=43"
      _StyleDefs(162) =   "Splits(1).Columns(12).Style:id=94,.parent=25"
      _StyleDefs(163) =   "Splits(1).Columns(12).HeadingStyle:id=91,.parent=26"
      _StyleDefs(164) =   "Splits(1).Columns(12).FooterStyle:id=92,.parent=27"
      _StyleDefs(165) =   "Splits(1).Columns(12).EditorStyle:id=93,.parent=43"
      _StyleDefs(166) =   "Splits(1).Columns(13).Style:id=114,.parent=25,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(167) =   "Splits(1).Columns(13).HeadingStyle:id=111,.parent=26"
      _StyleDefs(168) =   "Splits(1).Columns(13).FooterStyle:id=112,.parent=27"
      _StyleDefs(169) =   "Splits(1).Columns(13).EditorStyle:id=113,.parent=43"
      _StyleDefs(170) =   "Splits(1).Columns(14).Style:id=158,.parent=25"
      _StyleDefs(171) =   "Splits(1).Columns(14).HeadingStyle:id=155,.parent=26"
      _StyleDefs(172) =   "Splits(1).Columns(14).FooterStyle:id=156,.parent=27"
      _StyleDefs(173) =   "Splits(1).Columns(14).EditorStyle:id=157,.parent=43"
      _StyleDefs(174) =   "Named:id=33:Normal"
      _StyleDefs(175) =   ":id=33,.parent=0"
      _StyleDefs(176) =   "Named:id=34:Heading"
      _StyleDefs(177) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(178) =   ":id=34,.wraptext=-1"
      _StyleDefs(179) =   "Named:id=35:Footing"
      _StyleDefs(180) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(181) =   "Named:id=36:Selected"
      _StyleDefs(182) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(183) =   "Named:id=37:Caption"
      _StyleDefs(184) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(185) =   "Named:id=38:HighlightRow"
      _StyleDefs(186) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(187) =   "Named:id=39:EvenRow"
      _StyleDefs(188) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(189) =   "Named:id=40:OddRow"
      _StyleDefs(190) =   ":id=40,.parent=33"
      _StyleDefs(191) =   "Named:id=41:RecordSelector"
      _StyleDefs(192) =   ":id=41,.parent=34"
      _StyleDefs(193) =   "Named:id=42:FilterBar"
      _StyleDefs(194) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBDropDown tdbdTipoAcciones 
      Height          =   1155
      Left            =   3645
      TabIndex        =   10
      Top             =   2835
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   2037
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descripción"
      Columns(0).DataField=   "TAB_CDESCRIPCAMPO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "TAB_CCODIGO"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   0
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   -1  'True
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Arial"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF1EFEB&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.bgcolor=&HDDFFFE&"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
   Begin TDBNumber6Ctl.TDBNumber tdbNumber 
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      Top             =   3645
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManValores.frx":0F51
      Caption         =   "frmManValores.frx":0F71
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManValores.frx":0FD5
      Keys            =   "frmManValores.frx":0FF3
      Spin            =   "frmManValores.frx":102D
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
      EditMode        =   0
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
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumberNeg 
      Height          =   285
      Left            =   2025
      TabIndex        =   12
      Top             =   3735
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      Calculator      =   "frmManValores.frx":1055
      Caption         =   "frmManValores.frx":1075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManValores.frx":10D9
      Keys            =   "frmManValores.frx":10F7
      Spin            =   "frmManValores.frx":1131
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;-###,###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999
      MinValue        =   -999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   375
      Left            =   9630
      TabIndex        =   7
      Top             =   630
      Width           =   1380
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2434;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdVerificar 
      Height          =   375
      Left            =   3285
      TabIndex        =   3
      ToolTipText     =   " Verifica los importes"
      Top             =   630
      Width           =   1380
      Caption         =   " Verificar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManValores.frx":1159
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   375
      Left            =   225
      TabIndex        =   1
      ToolTipText     =   " Vuelve a cargar los datos almacenados "
      Top             =   630
      Width           =   1380
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManValores.frx":16F3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   1755
      TabIndex        =   2
      ToolTipText     =   "Grabar modificaciones"
      Top             =   630
      Width           =   1380
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManValores.frx":1C8D
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   4860
      TabIndex        =   4
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   630
      Width           =   1380
      Caption         =   " Insertar Mov."
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManValores.frx":2227
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   6435
      TabIndex        =   5
      ToolTipText     =   "Eliminar el movimientos seleccionado"
      Top             =   630
      Width           =   1380
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "frmManValores.frx":27C1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   8010
      TabIndex        =   6
      ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
      Top             =   630
      Width           =   1380
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2434;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   225
      TabIndex        =   9
      Top             =   225
      Width           =   780
   End
End
Attribute VB_Name = "frmManValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrCapital As New XArrayDB
Dim gsGrupo, sqlcombos As String
Dim lArrDetalle(14) As Variant
Dim rsTipoAcciones As ADODB.Recordset
Dim gsSalirControl As Boolean 'PARA EL CONTROL TDBNUMBER QUE ESTA ASOCIADA A LA GRILLA DEL DETALLE
Dim gsColumna As Integer
Const gsColorBloqueado = &HFFDBBB
Const NUM_COL = 15

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdEliminaItem_Click()
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If


   If CE(grdCapital.Columns(0).Value) = "" Then
      Mensajes "Seleccione una fila con datos"
      Exit Sub
   End If

   If MsgBox("Deseas eliminar la fila seleccionada", vbYesNo + vbQuestion) = vbYes Then
      On Error Resume Next
      lArrCapital.DeleteRows grdCapital.Bookmark
      DoEvents
      grdCapital.Update
      DoEvents
      grdCapital.ReBind
   End If
End Sub

Private Function CuentaFilas() As Integer
    Dim i As Integer
    Dim Contador As Integer
    Contador = 0
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 0)) <> "" Then
           Contador = Contador + 1
        End If
    Next i
    
    CuentaFilas = Contador
End Function

Private Sub cmdEliminarTodo_Click()

    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If

    If MsgBox("Deseas eliminar todas las filas de la lista", vbYesNo + vbQuestion) = vbYes Then
       lArrCapital.ReDim 0, 0, 0, NUM_COL ' filas
       lArrCapital.Clear
       DoEvents
       grdCapital.Update
       DoEvents
       grdCapital.ReBind
    End If
End Sub

Private Sub cmdGrabar_Click()
    Grabar
End Sub

Private Sub cmdInsertarItem_Click()
    Dim sql As String
    Dim rsVouchers As ADODB.Recordset
    
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If
    
    Call CerrarRecordSet(rsVouchers)
    
    sql = "spCn_RptFormato0308 'BUSCAR_VOUCHER', '" & gsEmpresa & "','" & gsAnio & "','" & tdbcMes.BoundText & "'"
    Call LlenarRecordSet(sql, rsVouchers)
    
    Dim Filas As Integer
    Filas = CuentaFilas 'lArrCapital.Count(1)
    
    On Error GoTo Siguiente
    If Filas = 1 And CE(lArrCapital(0, 0)) = "" Then Filas = 0
    
Siguiente:
    If Filas < 0 Then Filas = 0
    If Not rsVouchers Is Nothing Then
       Do While Not rsVouchers.EOF
          If rsVouchers.RecordCount <= 0 Then
             Mensajes "No se encontraron vouchers con las cuenta de Valores y entidad"
             Exit Sub
          End If
          
          If BuscaEntidad(CE(rsVouchers.Fields("ASE_NVOUCHER")), _
                          NE(rsVouchers.Fields("ASD_NITEM")), _
                          CE(rsVouchers.Fields("TEN_CTIPOENTIDAD")), _
                          CE(rsVouchers.Fields("ENT_CCODENTIDAD"))) = False Then
              
              lArrCapital.ReDim 0, Filas, 0, NUM_COL    ' filas
              
              lArrCapital(Filas, 0) = CE(rsVouchers.Fields("ASE_NVOUCHER"))  'voucher
              lArrCapital(Filas, 1) = CE(rsVouchers.Fields("ASD_NITEM")) 'item
              lArrCapital(Filas, 2) = CE(rsVouchers.Fields("TEN_CTIPOENTIDAD")) 'entidad
              lArrCapital(Filas, 3) = CE(rsVouchers.Fields("ENT_CCODENTIDAD")) 'codigo
              lArrCapital(Filas, 4) = CE(rsVouchers.Fields("ENT_CPERSONA")) 'apellidos
              lArrCapital(Filas, 5) = "" 'tipo de titulo
              lArrCapital(Filas, 6) = "" 'desc titulo
              lArrCapital(Filas, 7) = "0.00" 'valor nom unit
              lArrCapital(Filas, 8) = "0.00" 'cantidad
              
              If CE(rsVouchers.Fields("PROVTOTAL")) = "N" Then
                 lArrCapital(Filas, 9) = NE(rsVouchers.Fields("IMPORTE")) 'costo total
                 lArrCapital(Filas, 10) = "0.00" 'prov total
                 lArrCapital(Filas, 11) = "0.00" 'Otros Costos
                 lArrCapital(Filas, 12) = "0.00" 'Ajus. Val. Razo.
                 
                 If NE(rsVouchers.Fields("PER_CPERIODO")) = "00" And NE(rsVouchers.Fields("IMPORTE")) < 0 Then
                    lArrCapital(Filas, 9) = NE(rsVouchers.Fields("IMPORTE")) 'costo total
                    lArrCapital(Filas, 10) = "0.00" 'prov total
                    lArrCapital(Filas, 11) = "0.00" 'prov total
                    lArrCapital(Filas, 12) = "0.00" 'prov total
                 End If
              Else
                 lArrCapital(Filas, 7) = "" 'valor nom unit
                 lArrCapital(Filas, 8) = "" 'cantidad
                 lArrCapital(Filas, 9) = "0.00" 'costo total
                 lArrCapital(Filas, 10) = NE(rsVouchers.Fields("IMPORTE")) 'prov total
                 lArrCapital(Filas, 11) = "0.00" 'Otros Costos
                 lArrCapital(Filas, 12) = "0.00" 'Ajus. Razo.
              End If
              
              
              lArrCapital(Filas, 13) = (lArrCapital(Filas, 9) + lArrCapital(Filas, 11) + lArrCapital(Filas, 12)) - lArrCapital(Filas, 10)    'total neto
              lArrCapital(Filas, 14) = "0" 'flag
              lArrCapital(Filas, 15) = CE(rsVouchers.Fields("PROVTOTAL"))  'flag

              Filas = Filas + 1
          End If
            
          rsVouchers.MoveNext
       Loop
       grdCapital.ReBind
    Else
       Mensajes "No se encontraron vouchers con las cuenta de Valores y entidad"
    End If
        
End Sub

Private Function BuscaEntidad(Voucher As String, item As Integer, Tipo As String, Codigo As String) As Boolean
    BuscaEntidad = False
    Dim i As Integer
    On Error GoTo serror
    If (lArrCapital.Count(1) = 1 Or lArrCapital.Count(2) = 1) And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
       BuscaEntidad = False
       Exit Function
    End If
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 0)) = Voucher And _
           NE(lArrCapital(i, 1)) = item And _
           CE(lArrCapital(i, 2)) = Tipo And _
           CE(lArrCapital(i, 3)) = Codigo Then
           
           BuscaEntidad = True
           Exit For
        End If
    Next i
    Exit Function
serror:
    BuscaEntidad = False
End Function

Private Sub cmdRefresh_Click()
    cmdRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    
    GeneraArreglo
    DoEvents
    cmdRefresh.Enabled = True
    Screen.MousePointer = vbNormal
    
End Sub

Private Function BuscaImporte(ByRef rs As ADODB.Recordset, Voucher As String, item As Integer, Tipo As String, Codigo As String) As Double
    BuscaImporte = -1
    
    If Not rs Is Nothing Then
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
               If rs.Fields("ASE_NVOUCHER") = Voucher And _
                  rs.Fields("ASD_NITEM") = item And _
                  rs.Fields("TEN_CTIPOENTIDAD") = Tipo And _
                  rs.Fields("ENT_CCODENTIDAD") = Codigo Then
                  
                  BuscaImporte = NE(rs.Fields("IMPORTE"))
                  Exit Do
               End If
               rs.MoveNext
            Loop
        End If
    End If
End Function

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdVerificar_Click()

    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If


    If lArrCapital.Count(2) = 1 Then
        Exit Sub
    End If
    
    Dim sql As String
    Dim rsVouchers As ADODB.Recordset
    
    cmdVerificar.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Call CerrarRecordSet(rsVouchers)
    
    sql = "spCn_RptFormato0308 'VERIFICAR_VOUCHER', '" & gsEmpresa & "','" & gsAnio & "','" & tdbcMes.BoundText & "'"
    Call LlenarRecordSet(sql, rsVouchers)
    
    Dim i As Integer
    Dim Importe As Double
    For i = 0 To lArrCapital.Count(1) - 1
        Importe = BuscaImporte(rsVouchers, _
                               CE(lArrCapital(i, 0)), _
                               NE(lArrCapital(i, 1)), _
                               CE(lArrCapital(i, 2)), _
                               CE(lArrCapital(i, 3)))
        
        If Importe = NE(lArrCapital(i, 11)) Then lArrCapital(i, 12) = "0"
        If Importe <> NE(lArrCapital(i, 11)) Then lArrCapital(i, 12) = "1"
        If Importe = -1 Then lArrCapital(i, 12) = "2"
        If CE(lArrCapital(i, 0)) = "" Then lArrCapital(i, 12) = "0"
    Next i

    grdCapital.Refresh
    
    cmdVerificar.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
                If MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar") = vbYes Then
                    Unload Me
                End If

        Case 115: If cmdGrabar.Enabled Then cmdGrabar_Click
        'Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        'Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select


End Sub

Private Sub CargarCombos()
    Dim sqlcombos  As String
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    DoEvents
    tdbcMes.BoundText = gsPeriodo
    '---------------------------------------------------------
    CerrarRecordSet rsTipoAcciones
    
    sqlcombos = "SELECT TAB_CDESCRIPCAMPO , TAB_CCODIGO " & _
                "FROM TABLA WHERE TAB_CTABLA='069' AND EMP_CCODIGO='" & gsEmpresa & "' " & _
                "ORDER BY TAB_CDESCRIPCAMPO "
    
    Call LlenarRecordSet(sqlcombos, rsTipoAcciones)
    Set tdbdTipoAcciones.DataSource = rsTipoAcciones


End Sub

Private Sub SumarTotales()
    Dim i As Integer
    Dim iFila As Integer
    
    
    Dim s_Importe As Double, s_valNom As Double
    Dim s_AccSus As Double, s_AccPag As Double
    Dim s_NroAcc As Double
    
    On Error GoTo serror
    iFila = lArrCapital.Count(1)
    
    For i = 0 To iFila - 1
        s_Importe = s_Importe + NE(lArrCapital.Value(i, 7))
        s_valNom = s_valNom + NE(lArrCapital.Value(i, 8))
        s_AccSus = s_AccSus + NE(lArrCapital.Value(i, 9))
        s_AccPag = s_AccPag + NE(lArrCapital.Value(i, 10))
        s_NroAcc = s_NroAcc + NE(lArrCapital.Value(i, 11))
    Next i

    grdCapital.Columns(7).FooterText = Format(s_Importe, "###,###,##0.00")
    grdCapital.Columns(8).FooterText = Format(s_valNom, "###,###,##0.00")
    grdCapital.Columns(9).FooterText = Format(s_AccSus, "###,###,##0.00")
    grdCapital.Columns(10).FooterText = Format(s_AccPag, "###,###,##0.00")
    grdCapital.Columns(11).FooterText = Format(s_NroAcc, "###,###,##0.00")

    
    Exit Sub
    
serror:
    
     s_Importe = 0
     s_valNom = 0
     s_AccSus = 0
     s_AccPag = 0
     s_NroAcc = 0

    grdCapital.Columns(7).FooterText = Format(s_Importe, "###,###,##0.00")
    grdCapital.Columns(8).FooterText = Format(s_valNom, "###,###,##0.00")
    grdCapital.Columns(9).FooterText = Format(s_AccSus, "###,###,##0.00")
    grdCapital.Columns(10).FooterText = Format(s_AccPag, "###,###,##0.00")
    grdCapital.Columns(11).FooterText = Format(s_NroAcc, "###,###,##0.00")

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    
    Me.Height = 6990
    Me.Width = 11550
    
    grdCapital.Splits(0).MarqueeStyle = dbgHighlightRow
    grdCapital.HighlightRowStyle = "HighlightRow"
    
    
    GeneraArreglo
    
    DoEvents
    grdCapital.FetchRowStyle = True
    grdCapital.Columns(7).FetchStyle = True
    grdCapital.Columns(8).FetchStyle = True
    
    SumarTotales
    CargarCombos
    
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabar.Enabled = False
        cmdInsertarItem.Enabled = False
        cmdEliminaItem.Enabled = False
        cmdEliminarTodo.Enabled = False
        grdCapital.Splits(1).Locked = True
    Else
        cmdGrabar.Enabled = True
        cmdInsertarItem.Enabled = True
        cmdEliminaItem.Enabled = True
        cmdEliminarTodo.Enabled = True
        grdCapital.Splits(1).Locked = False
    End If
    
    
End Sub

Private Sub GeneraArreglo()
    Dim sql As String
    
    On Local Error GoTo ErrorEjecucion
    
    Set lArrCapital = Nothing
    Set grdCapital.Array = lArrCapital
    grdCapital.ReBind
    
    If tdbcMes.Text = "" Then
'        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If
    
    
    sql = "spCn_RptFormato0308 'BUSCARTODOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "'"
    Call GridArreglo(lArrCapital, grdCapital, sql)
    
    grdCapital.Splits(1).ScrollBars = dbgBoth
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        'fraEstructura.Height = Me.Height - 1900 + 250
        'fraEstructura.Width = Me.Width - 150
        grdCapital.Height = Me.Height - 2300 + 150
        grdCapital.Width = Me.Width - 400
        grdCapital.Splits(1).ScrollBars = dbgNone
        grdCapital.Splits(1).ScrollBars = dbgAutomatic
        fraLeyenda.Width = grdCapital.Width
    End If
    
    Exit Sub
    
serror:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub grdCapital_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 8 Then
        If NE(grdCapital.Columns(11).Value) < 0 Then
           grdCapital.Columns(ColIndex) = Abs(grdCapital.Columns(ColIndex).Value) * -1
        Else
           grdCapital.Columns(ColIndex) = Abs(grdCapital.Columns(ColIndex).Value)
        End If
    End If
    
End Sub

Private Sub grdCapital_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

'    If CE(lArrCapital(0, grdCapital.Bookmark)) = "" Then
'       Cancel = 1
'    End If
        
'    If ColIndex = 7 Or ColIndex = 8 Then
'        If CE(lArrCapital(grdCapital.Bookmark, 13)) = "S" Then
'           Cancel = 1
'        End If
'    End If
End Sub

Private Sub UpdateGrilla()
    On Error Resume Next
    DoEvents
    grdCapital.Update
    DoEvents
End Sub


Private Sub grdCapital_BeforeRowColChange(Cancel As Integer)
    If gsColumna = 6 Or gsColumna = 7 Or gsColumna = 8 Then
        Dim Producto As Double
        UpdateGrilla
        
        Producto = NE(lArrCapital(grdCapital.Bookmark, 7)) * NE(lArrCapital(grdCapital.Bookmark, 8))
        
        If Producto <> NE(lArrCapital(grdCapital.Bookmark, 11)) And Producto <> 0 Then
            lArrCapital(grdCapital.Bookmark, 12) = "3"
            grdCapital.Refresh
        Else
            lArrCapital(grdCapital.Bookmark, 12) = ""
            grdCapital.Refresh
        End If
        
'        If CE(lArrCapital(grdCapital.Bookmark, 13)) = "S" And (gsColumna = 7 Or gsColumna = 8) Then
'            TDBNumberNeg.BackColor = gsColorBloqueado
'        Else
'            TDBNumberNeg.BackColor = gsColorActivado
'        End If
        
    End If
End Sub

Private Sub grdCapital_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
    On Error GoTo serror
'    If Split = 1 Then
'        If Col = 7 Or Col = 8 Then
'            If lArrCapital.Count(2) > 1 Then
'                If CE(lArrCapital(Bookmark, 13)) = "S" Then
'                    CellStyle.BackColor = gsColorBloqueado
'                End If
'            End If
'        End If
'    End If
    Exit Sub
serror:
End Sub

Private Sub grdCapital_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    If lArrCapital Is Nothing Or IsNull(grdCapital.Bookmark) Then
        Exit Sub
    End If
    
    On Error GoTo serror
    
    If lArrCapital(Bookmark, 12) = "2" Then
        RowStyle.BackColor = &HFF&
        RowStyle.ForeColor = &HFFFF&
    ElseIf lArrCapital(Bookmark, 12) = "1" Then
        RowStyle.BackColor = gsColorDesactProv
    ElseIf lArrCapital(Bookmark, 12) = "3" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    
    Exit Sub
serror:
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
    Select Case lControl
           Case "CuentasFilt"
                grdCapital.Columns(grdCapital.Col).Value = param0
           
    End Select
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False
    
    If lArrCapital.Count(1) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
       ValidaCampos = True
       Exit Function
    End If
    
    Dim i As Integer
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 7)) = "" Then
            lArrCapital(i, 7) = 0
        End If
        
        If CE(lArrCapital(i, 8)) = "" Then
            lArrCapital(i, 8) = 0
        End If
    
        If CE(lArrCapital(i, 5)) = "" Then
           Mensajes "Ingrese un TIPO DE TITULO en el Voucher " & lArrCapital(i, 0) & " de: " & Salto(2) & lArrCapital(i, 4)
           grdCapital.Bookmark = i
           grdCapital.Col = 5
           pSetFocus grdCapital
           Exit Function
        End If
        
'        If NE(lArrCapital(i, 7)) = 0 And CE(lArrCapital(i, 13)) <> "S" Then
'           Mensajes "Ingrese el VALOR NOMINAL unitario en el Voucher " & lArrCapital(i, 0) & " de " & lArrCapital(i, 4)
'           grdCapital.Bookmark = i
'           grdCapital.Col = 7
'           pSetFocus grdCapital
'           Exit Function
'        End If
'
'
'        If NE(lArrCapital(i, 8)) = 0 And CE(lArrCapital(i, 13)) <> "S" Then
'           Mensajes "Ingrese la CANTIDAD en el voucher " & lArrCapital(i, 0) & " de " & lArrCapital(i, 4)
'           grdCapital.Bookmark = i
'           grdCapital.Col = 8
'           pSetFocus grdCapital
'           Exit Function
'        End If
        
    
        
        If NE(lArrCapital(i, 11)) = 0 Then
           Mensajes "Ingrese minimo una accion en el Voucher " & lArrCapital(i, 0) & " de " & lArrCapital(i, 4)
           grdCapital.Bookmark = i
           grdCapital.Col = 9
           pSetFocus grdCapital
           Exit Function
        End If
        
    Next i
    ValidaCampos = True
End Function


Private Sub Grabar()
    On Error GoTo serror
    
    If tdbcMes.Text = "" Then
        Mensajes "Seleccione el periodo"
        pSetFocus tdbcMes
        Exit Sub
    End If
    
    UpdateGrilla
    
    If lArrCapital.Count(2) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
        Exit Sub
    End If

    If ValidaCampos = False Then Exit Sub
    
    'UpdateGrilla

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
    Set clsMante = New clsMantoTablas
    grdCapital.Bookmark = grdCapital.Bookmark

    Dim lArrDet(3) As Variant
    lArrDet(0) = "ELIMINAR"
    lArrDet(1) = gsEmpresa
    lArrDet(2) = gsAnio
    lArrDet(3) = tdbcMes.BoundText
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_RptFormato0308", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    For i = 0 To lArrCapital.Count(1) - 1
        If CE(lArrCapital(i, 0)) <> "" Then
            
                If CargaArregloDet(i) = True Then
                    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_RptFormato0308", lArrDetalle(), False) = False Then
                        Screen.MousePointer = vbNormal
                        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                        Exit Sub
                    End If
                End If
            
        End If
    Next

    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal

    Set clsMante = Nothing
    Call GeneraArreglo
    
    DoEvents
    
    cmdRefresh_Click
    
    DoEvents
    Mensajes "Se ha grabado con exito ...", vbInformation
   
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Private Function CargaArregloDet(item As Integer) As Boolean
    CargaArregloDet = True
    
    lArrDetalle(0) = "INSERTAR"
    lArrDetalle(1) = gsEmpresa
    lArrDetalle(2) = gsAnio
    lArrDetalle(3) = tdbcMes.BoundText
    lArrDetalle(4) = CE(lArrCapital(item, 0)) 'voucher
    lArrDetalle(5) = NE(lArrCapital(item, 1)) 'item
    lArrDetalle(6) = CE(lArrCapital(item, 2)) 'tipo entidad
    lArrDetalle(7) = CE(lArrCapital(item, 3)) 'cod entidad
    lArrDetalle(8) = CE(lArrCapital(item, 5)) 'cod titulo
    lArrDetalle(9) = CE(lArrCapital(item, 6)) 'desc titulo
    lArrDetalle(10) = NE(lArrCapital(item, 7)) 'valor nom unit
    lArrDetalle(11) = NE(lArrCapital(item, 8)) 'cantidad
    lArrDetalle(12) = NE(lArrCapital(item, 9)) 'costo total
    lArrDetalle(13) = NE(lArrCapital(item, 10)) 'prov total
    lArrDetalle(14) = NE(lArrCapital(item, 11)) 'total neto

End Function

Private Sub Imprimir()
    Screen.MousePointer = vbHourglass

    Dim matriz_fecha(4) As Variant
    Dim Tipo As String

    matriz_fecha(0) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(1) = "@RUC;" & gsRUC & ";True"
    matriz_fecha(2) = "@Tipo;BUSCARTODOS;True"
    matriz_fecha(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(4) = "@Pan_cAnio;" & gsAnio & ";True"

    Dim formulas(0) As Variant
    'AbreReporteParam gsDSN, Me, rutaReportes & "RptPatrimonioNeto.rpt", crptToWindow, "Reporte de Patrimonio Neto", "", matriz_fecha(), Formulas()
    Screen.MousePointer = vbDefault

End Sub

Private Sub grdCapital_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And gsColumna >= 8 Then
       grdCapital.Col = 6
       grdCapital.Bookmark = grdCapital.Bookmark + 1
       pSetFocus grdCapital
    
    End If
End Sub

Private Sub grdCapital_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If lArrCapital.Count(2) = 1 And grdCapital.Bookmark = 0 And CE(grdCapital.Columns(0)) = "" Then
        Exit Sub
    End If

    gsSalirControl = False
    
    grdCapital.Columns(11).Value = NE(grdCapital.Columns(9).Value) + NE(grdCapital.Columns(10).Value)
    
'    If lArrCapital.Count(2) > 1 Then
'        If CE(lArrCapital(grdCapital.Bookmark, 13)) = "S" And (gsColumna = 7 Or gsColumna = 8) Then
'            TDBNumberNeg.BackColor = gsColorBloqueado
'        Else
'            TDBNumberNeg.BackColor = gsColorActivado
'        End If
'    End If
    
    grdCapital.Update
    gsColumna = grdCapital.Col
    
    SumarTotales
    
End Sub

Private Sub tdbcMes_ItemChange()
    cmdRefresh_Click
End Sub


Private Sub tdbdTipoAcciones_DropDownClose()
    grdCapital.Columns(5) = tdbdTipoAcciones.Columns(1).Value
    grdCapital.Columns(6) = tdbdTipoAcciones.Columns(0).Value
    DoEvents
    
    grdCapital.RefetchRow
    pSetFocus grdCapital
    DoEvents

End Sub

Private Sub tdbNumber_GotFocus()
    On Error Resume Next
    TDBNumber.Value = Abs(NE(lArrCapital(grdCapital.Bookmark, grdCapital.Col)))
End Sub

Private Sub tdbNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       ControlAbs TDBNumber
       pSetFocus grdCapital
    End If
End Sub

Private Sub tdbNumber_KeyPress(KeyAscii As Integer)
    If gsSalirControl = False Then
        gsSalirControl = True
        TDBNumber = "0.00"
    End If
End Sub

