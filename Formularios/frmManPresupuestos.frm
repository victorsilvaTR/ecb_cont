VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManPresupuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Presupuesto Anual"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   Icon            =   "frmManPresupuestos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   10920
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   5310
      TabIndex        =   13
      Top             =   1305
      Width           =   5550
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro Costo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   4275
         TabIndex        =   16
         Top             =   315
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   2
         Left            =   3735
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub.Titulo C.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   15
         Top             =   315
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   1
         Left            =   1800
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo C.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   675
         TabIndex        =   14
         Top             =   315
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   0
         Left            =   135
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1485
      Left            =   135
      ScaleHeight     =   1425
      ScaleWidth      =   2580
      TabIndex        =   7
      Top             =   450
      Width           =   2640
      Begin VB.Image Image1 
         Height          =   3375
         Left            =   60
         Picture         =   "frmManPresupuestos.frx":0ECA
         Top             =   -1050
         Width           =   2490
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   2985
      TabIndex        =   1
      Top             =   360
      Width           =   2205
      Begin VB.OptionButton optIngresos 
         Caption         =   "INGRESOS"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   810
         Width           =   1275
      End
      Begin VB.OptionButton optEgresos 
         Caption         =   "GASTOS"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   540
         Value           =   -1  'True
         Width           =   1170
      End
      Begin MSForms.CommandButton cmdListar 
         Height          =   375
         Left            =   225
         TabIndex        =   17
         ToolTipText     =   "Cargar nueva Configuración"
         Top             =   1125
         Width           =   1575
         Caption         =   " Listar"
         PicturePosition =   327683
         Size            =   "2778;661"
         Picture         =   "frmManPresupuestos.frx":9A57
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblAnio 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "AÑO :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   270
         TabIndex        =   6
         Top             =   180
         Width           =   1650
      End
   End
   Begin TrueOleDBGrid70.TDBGrid grdEgresos 
      Height          =   4110
      Left            =   105
      TabIndex        =   0
      Top             =   2010
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   7250
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nivel"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Total"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Enero"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Febrero"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Marzo"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Abril"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Mayo"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Junio"
      Columns(9).DataField=   ""
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Julio"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Agosto"
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Setiembre"
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "Standard"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Octubre"
      Columns(13).DataField=   ""
      Columns(13).NumberFormat=   "Standard"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Noviembre"
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Diciembre"
      Columns(15).DataField=   ""
      Columns(15).NumberFormat=   "Standard"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "TC"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Moneda"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   18
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).SizeMode=   2
      Splits(0).Size  =   3
      Splits(0).Size.vt=   2
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=18"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=4948"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=4868"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=265"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=185"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2064"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=1852"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(37)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(38)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(39)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(40)=   "Column(5).Width=1852"
      Splits(0)._ColumnProps(41)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(43)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(45)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(46)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(47)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(48)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(49)=   "Column(6).Width=1852"
      Splits(0)._ColumnProps(50)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(52)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(54)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(55)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(56)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(57)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(58)=   "Column(7).Width=1852"
      Splits(0)._ColumnProps(59)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(7)._WidthInPix=1773"
      Splits(0)._ColumnProps(61)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(62)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(63)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(64)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(65)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(67)=   "Column(8).Width=1852"
      Splits(0)._ColumnProps(68)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(8)._WidthInPix=1773"
      Splits(0)._ColumnProps(70)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(72)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(73)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(74)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(75)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(76)=   "Column(9).Width=1852"
      Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=1773"
      Splits(0)._ColumnProps(79)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(81)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(82)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(83)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(84)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(85)=   "Column(10).Width=1852"
      Splits(0)._ColumnProps(86)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(10)._WidthInPix=1773"
      Splits(0)._ColumnProps(88)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(89)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(90)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(91)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(92)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(93)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(94)=   "Column(11).Width=1852"
      Splits(0)._ColumnProps(95)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(11)._WidthInPix=1773"
      Splits(0)._ColumnProps(97)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(98)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(99)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(100)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(101)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(102)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(103)=   "Column(12).Width=1852"
      Splits(0)._ColumnProps(104)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(12)._WidthInPix=1773"
      Splits(0)._ColumnProps(106)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(107)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(108)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(109)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(110)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(111)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(112)=   "Column(13).Width=1852"
      Splits(0)._ColumnProps(113)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(114)=   "Column(13)._WidthInPix=1773"
      Splits(0)._ColumnProps(115)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(116)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(117)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(118)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(119)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(120)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(121)=   "Column(14).Width=1852"
      Splits(0)._ColumnProps(122)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(14)._WidthInPix=1773"
      Splits(0)._ColumnProps(124)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(125)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(126)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(127)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(128)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(129)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(130)=   "Column(15).Width=1852"
      Splits(0)._ColumnProps(131)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(132)=   "Column(15)._WidthInPix=1773"
      Splits(0)._ColumnProps(133)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(134)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(135)=   "Column(15)._ColStyle=514"
      Splits(0)._ColumnProps(136)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(137)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(138)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(139)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(140)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(142)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(143)=   "Column(16).AllowSizing=0"
      Splits(0)._ColumnProps(144)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(145)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(146)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(147)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(148)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(149)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(151)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(152)=   "Column(17).AllowSizing=0"
      Splits(0)._ColumnProps(153)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(154)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(155)=   "Column(17).AllowFocus=0"
      Splits(0)._ColumnProps(156)=   "Column(17).Order=18"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=18"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=4075"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=3995"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=2143"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=2064"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=514"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=1852"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=1773"
      Splits(1)._ColumnProps(36)=   "Column(4)._ColStyle=514"
      Splits(1)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(38)=   "Column(5).Width=1852"
      Splits(1)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(40)=   "Column(5)._WidthInPix=1773"
      Splits(1)._ColumnProps(41)=   "Column(5)._ColStyle=514"
      Splits(1)._ColumnProps(42)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(43)=   "Column(6).Width=1852"
      Splits(1)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(45)=   "Column(6)._WidthInPix=1773"
      Splits(1)._ColumnProps(46)=   "Column(6)._ColStyle=514"
      Splits(1)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(48)=   "Column(7).Width=1852"
      Splits(1)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(50)=   "Column(7)._WidthInPix=1773"
      Splits(1)._ColumnProps(51)=   "Column(7)._ColStyle=514"
      Splits(1)._ColumnProps(52)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(53)=   "Column(8).Width=1852"
      Splits(1)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(55)=   "Column(8)._WidthInPix=1773"
      Splits(1)._ColumnProps(56)=   "Column(8)._ColStyle=514"
      Splits(1)._ColumnProps(57)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(58)=   "Column(9).Width=1852"
      Splits(1)._ColumnProps(59)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(60)=   "Column(9)._WidthInPix=1773"
      Splits(1)._ColumnProps(61)=   "Column(9)._ColStyle=514"
      Splits(1)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(63)=   "Column(10).Width=1852"
      Splits(1)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(10)._WidthInPix=1773"
      Splits(1)._ColumnProps(66)=   "Column(10)._ColStyle=514"
      Splits(1)._ColumnProps(67)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(68)=   "Column(11).Width=1852"
      Splits(1)._ColumnProps(69)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(70)=   "Column(11)._WidthInPix=1773"
      Splits(1)._ColumnProps(71)=   "Column(11)._ColStyle=514"
      Splits(1)._ColumnProps(72)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(73)=   "Column(12).Width=1852"
      Splits(1)._ColumnProps(74)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(75)=   "Column(12)._WidthInPix=1773"
      Splits(1)._ColumnProps(76)=   "Column(12)._ColStyle=514"
      Splits(1)._ColumnProps(77)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(78)=   "Column(13).Width=1852"
      Splits(1)._ColumnProps(79)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(80)=   "Column(13)._WidthInPix=1773"
      Splits(1)._ColumnProps(81)=   "Column(13)._ColStyle=514"
      Splits(1)._ColumnProps(82)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(83)=   "Column(14).Width=1852"
      Splits(1)._ColumnProps(84)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(85)=   "Column(14)._WidthInPix=1773"
      Splits(1)._ColumnProps(86)=   "Column(14)._ColStyle=514"
      Splits(1)._ColumnProps(87)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(88)=   "Column(15).Width=1852"
      Splits(1)._ColumnProps(89)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(90)=   "Column(15)._WidthInPix=1773"
      Splits(1)._ColumnProps(91)=   "Column(15)._ColStyle=514"
      Splits(1)._ColumnProps(92)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(93)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(94)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(95)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(96)=   "Column(16)._ColStyle=516"
      Splits(1)._ColumnProps(97)=   "Column(16).Visible=0"
      Splits(1)._ColumnProps(98)=   "Column(16).AllowFocus=0"
      Splits(1)._ColumnProps(99)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(100)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(101)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(102)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(103)=   "Column(17)._ColStyle=516"
      Splits(1)._ColumnProps(104)=   "Column(17).Visible=0"
      Splits(1)._ColumnProps(105)=   "Column(17).AllowFocus=0"
      Splits(1)._ColumnProps(106)=   "Column(17).Order=18"
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
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HF1EFEB&"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=98,.parent=13,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=82,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=86,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=90,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=94,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=102,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=106,.parent=13,.bgcolor=&HFFFFFF&"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(109) =   "Splits(1).Style:id=107,.parent=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(110) =   "Splits(1).CaptionStyle:id=116,.parent=4"
      _StyleDefs(111) =   "Splits(1).HeadingStyle:id=108,.parent=2"
      _StyleDefs(112) =   "Splits(1).FooterStyle:id=109,.parent=3"
      _StyleDefs(113) =   "Splits(1).InactiveStyle:id=110,.parent=5"
      _StyleDefs(114) =   "Splits(1).SelectedStyle:id=112,.parent=6"
      _StyleDefs(115) =   "Splits(1).EditorStyle:id=111,.parent=7"
      _StyleDefs(116) =   "Splits(1).HighlightRowStyle:id=113,.parent=8"
      _StyleDefs(117) =   "Splits(1).EvenRowStyle:id=114,.parent=9"
      _StyleDefs(118) =   "Splits(1).OddRowStyle:id=115,.parent=10"
      _StyleDefs(119) =   "Splits(1).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(120) =   "Splits(1).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(121) =   "Splits(1).Columns(0).Style:id=122,.parent=107,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(122) =   "Splits(1).Columns(0).HeadingStyle:id=119,.parent=108"
      _StyleDefs(123) =   "Splits(1).Columns(0).FooterStyle:id=120,.parent=109"
      _StyleDefs(124) =   "Splits(1).Columns(0).EditorStyle:id=121,.parent=111"
      _StyleDefs(125) =   "Splits(1).Columns(1).Style:id=126,.parent=107,.bgcolor=&HFFFFFF&,.locked=0"
      _StyleDefs(126) =   "Splits(1).Columns(1).HeadingStyle:id=123,.parent=108"
      _StyleDefs(127) =   "Splits(1).Columns(1).FooterStyle:id=124,.parent=109"
      _StyleDefs(128) =   "Splits(1).Columns(1).EditorStyle:id=125,.parent=111"
      _StyleDefs(129) =   "Splits(1).Columns(2).Style:id=130,.parent=107,.bgcolor=&HFFFFFF&"
      _StyleDefs(130) =   "Splits(1).Columns(2).HeadingStyle:id=127,.parent=108"
      _StyleDefs(131) =   "Splits(1).Columns(2).FooterStyle:id=128,.parent=109"
      _StyleDefs(132) =   "Splits(1).Columns(2).EditorStyle:id=129,.parent=111"
      _StyleDefs(133) =   "Splits(1).Columns(3).Style:id=134,.parent=107,.alignment=1,.bgcolor=&HF1EFEB&"
      _StyleDefs(134) =   "Splits(1).Columns(3).HeadingStyle:id=131,.parent=108"
      _StyleDefs(135) =   "Splits(1).Columns(3).FooterStyle:id=132,.parent=109"
      _StyleDefs(136) =   "Splits(1).Columns(3).EditorStyle:id=133,.parent=111"
      _StyleDefs(137) =   "Splits(1).Columns(4).Style:id=138,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(138) =   "Splits(1).Columns(4).HeadingStyle:id=135,.parent=108"
      _StyleDefs(139) =   "Splits(1).Columns(4).FooterStyle:id=136,.parent=109"
      _StyleDefs(140) =   "Splits(1).Columns(4).EditorStyle:id=137,.parent=111"
      _StyleDefs(141) =   "Splits(1).Columns(5).Style:id=142,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(142) =   "Splits(1).Columns(5).HeadingStyle:id=139,.parent=108"
      _StyleDefs(143) =   "Splits(1).Columns(5).FooterStyle:id=140,.parent=109"
      _StyleDefs(144) =   "Splits(1).Columns(5).EditorStyle:id=141,.parent=111"
      _StyleDefs(145) =   "Splits(1).Columns(6).Style:id=146,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(146) =   "Splits(1).Columns(6).HeadingStyle:id=143,.parent=108"
      _StyleDefs(147) =   "Splits(1).Columns(6).FooterStyle:id=144,.parent=109"
      _StyleDefs(148) =   "Splits(1).Columns(6).EditorStyle:id=145,.parent=111"
      _StyleDefs(149) =   "Splits(1).Columns(7).Style:id=150,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(150) =   "Splits(1).Columns(7).HeadingStyle:id=147,.parent=108"
      _StyleDefs(151) =   "Splits(1).Columns(7).FooterStyle:id=148,.parent=109"
      _StyleDefs(152) =   "Splits(1).Columns(7).EditorStyle:id=149,.parent=111"
      _StyleDefs(153) =   "Splits(1).Columns(8).Style:id=154,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(154) =   "Splits(1).Columns(8).HeadingStyle:id=151,.parent=108"
      _StyleDefs(155) =   "Splits(1).Columns(8).FooterStyle:id=152,.parent=109"
      _StyleDefs(156) =   "Splits(1).Columns(8).EditorStyle:id=153,.parent=111"
      _StyleDefs(157) =   "Splits(1).Columns(9).Style:id=158,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(158) =   "Splits(1).Columns(9).HeadingStyle:id=155,.parent=108"
      _StyleDefs(159) =   "Splits(1).Columns(9).FooterStyle:id=156,.parent=109"
      _StyleDefs(160) =   "Splits(1).Columns(9).EditorStyle:id=157,.parent=111"
      _StyleDefs(161) =   "Splits(1).Columns(10).Style:id=162,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(162) =   "Splits(1).Columns(10).HeadingStyle:id=159,.parent=108"
      _StyleDefs(163) =   "Splits(1).Columns(10).FooterStyle:id=160,.parent=109"
      _StyleDefs(164) =   "Splits(1).Columns(10).EditorStyle:id=161,.parent=111"
      _StyleDefs(165) =   "Splits(1).Columns(11).Style:id=166,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(166) =   "Splits(1).Columns(11).HeadingStyle:id=163,.parent=108"
      _StyleDefs(167) =   "Splits(1).Columns(11).FooterStyle:id=164,.parent=109"
      _StyleDefs(168) =   "Splits(1).Columns(11).EditorStyle:id=165,.parent=111"
      _StyleDefs(169) =   "Splits(1).Columns(12).Style:id=170,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(170) =   "Splits(1).Columns(12).HeadingStyle:id=167,.parent=108"
      _StyleDefs(171) =   "Splits(1).Columns(12).FooterStyle:id=168,.parent=109"
      _StyleDefs(172) =   "Splits(1).Columns(12).EditorStyle:id=169,.parent=111"
      _StyleDefs(173) =   "Splits(1).Columns(13).Style:id=174,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(174) =   "Splits(1).Columns(13).HeadingStyle:id=171,.parent=108"
      _StyleDefs(175) =   "Splits(1).Columns(13).FooterStyle:id=172,.parent=109"
      _StyleDefs(176) =   "Splits(1).Columns(13).EditorStyle:id=173,.parent=111"
      _StyleDefs(177) =   "Splits(1).Columns(14).Style:id=178,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(178) =   "Splits(1).Columns(14).HeadingStyle:id=175,.parent=108"
      _StyleDefs(179) =   "Splits(1).Columns(14).FooterStyle:id=176,.parent=109"
      _StyleDefs(180) =   "Splits(1).Columns(14).EditorStyle:id=177,.parent=111"
      _StyleDefs(181) =   "Splits(1).Columns(15).Style:id=182,.parent=107,.alignment=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(182) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=108"
      _StyleDefs(183) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=109"
      _StyleDefs(184) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=111"
      _StyleDefs(185) =   "Splits(1).Columns(16).Style:id=186,.parent=107,.bgcolor=&HFFFFFF&"
      _StyleDefs(186) =   "Splits(1).Columns(16).HeadingStyle:id=183,.parent=108"
      _StyleDefs(187) =   "Splits(1).Columns(16).FooterStyle:id=184,.parent=109"
      _StyleDefs(188) =   "Splits(1).Columns(16).EditorStyle:id=185,.parent=111"
      _StyleDefs(189) =   "Splits(1).Columns(17).Style:id=190,.parent=107,.bgcolor=&HFFFFFF&"
      _StyleDefs(190) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=108"
      _StyleDefs(191) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=109"
      _StyleDefs(192) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=111"
      _StyleDefs(193) =   "Named:id=33:Normal"
      _StyleDefs(194) =   ":id=33,.parent=0"
      _StyleDefs(195) =   "Named:id=34:Heading"
      _StyleDefs(196) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(197) =   ":id=34,.wraptext=-1"
      _StyleDefs(198) =   "Named:id=35:Footing"
      _StyleDefs(199) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(200) =   "Named:id=36:Selected"
      _StyleDefs(201) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(202) =   "Named:id=37:Caption"
      _StyleDefs(203) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(204) =   "Named:id=38:HighlightRow"
      _StyleDefs(205) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(206) =   "Named:id=39:EvenRow"
      _StyleDefs(207) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(208) =   "Named:id=40:OddRow"
      _StyleDefs(209) =   ":id=40,.parent=33"
      _StyleDefs(210) =   "Named:id=41:RecordSelector"
      _StyleDefs(211) =   ":id=41,.parent=34"
      _StyleDefs(212) =   "Named:id=42:FilterBar"
      _StyleDefs(213) =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   360
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
            Picture         =   "frmManPresupuestos.frx":9FF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":A3CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":A7A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":AB7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":AF59
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":B333
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":B70D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":BAE7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkAcumular 
      Alignment       =   1  'Right Justify
      Caption         =   "Acumular Totales"
      Height          =   195
      Left            =   6615
      TabIndex        =   8
      Top             =   3195
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1665
   End
   Begin TrueOleDBList70.TDBCombo tdbcMoneda 
      Height          =   300
      Left            =   6705
      TabIndex        =   11
      Top             =   450
      Width           =   2670
      _ExtentX        =   4710
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
      Splits(0)._ColumnProps(7)=   "Column(1).Width=370"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=291"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1376"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1296"
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
      _PropDict       =   $"frmManPresupuestos.frx":CB01
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
            Picture         =   "frmManPresupuestos.frx":CB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":CCE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":CE3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":CF96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":D0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":D24A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":D3A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":D4FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":D658
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
            Picture         =   "frmManPresupuestos.frx":D7B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":DD4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":E2E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":E880
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":EE1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":F3B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":F94E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":FEE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPresupuestos.frx":10482
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   19
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin MSForms.CommandButton cmdSalir 
      Height          =   390
      Left            =   9450
      TabIndex        =   18
      Top             =   900
      Width           =   1290
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2275;688"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOrdenar 
      Height          =   390
      Left            =   8100
      TabIndex        =   12
      Top             =   900
      Width           =   1290
      Caption         =   " Ordenar Item"
      PicturePosition =   327683
      Size            =   "2275;688"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   390
      Left            =   6750
      TabIndex        =   10
      Top             =   900
      Width           =   1290
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2275;688"
      Picture         =   "frmManPresupuestos.frx":10A1C
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdActualiza 
      Height          =   390
      Left            =   5400
      TabIndex        =   9
      Top             =   900
      Width           =   1290
      Caption         =   " Insertar Item"
      PicturePosition =   327683
      Size            =   "2275;688"
      Picture         =   "frmManPresupuestos.frx":10FB6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "MONEDA : "
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
      Left            =   5355
      TabIndex        =   5
      Top             =   495
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   45
      Index           =   0
      Left            =   8010
      TabIndex        =   2
      Top             =   4230
      Width           =   300
   End
End
Attribute VB_Name = "frmManPresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrPspto As New XArrayDB
Dim valorPres As Double
Dim Fila As Integer
Dim lTipoPres As String
Dim lArrDet() As Variant
Dim lControl As String
Dim nFilas As Integer
Dim TCMensual(12) As Double
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Sub cmdActualiza_Click()
    Call LlamaBuscar(frmBuscador, "CentroCosto", lControl, "CentroCostoPres", Me, gsPeriodo)
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Dim i As Integer
    Dim psql As String
    Dim Valor As String
    Dim Titulo As String
    Dim STitulo As String
    
    Titulo = Left(param2, 1)
    STitulo = Right(param2, 1)
    
    Valor = True
    Select Case lControl
            Case "CentroCostoPres"
                ' *** Ver si centro de costo no esta en el Grid
                For i = 0 To lArrPspto.Count(1) - 1
                    If Trim(lArrPspto(i, 0)) = Trim(param0) Then
                        Valor = False
                        Exit For
                    End If
                Next
                ' ***
                Dim Fila As Integer
                Dim Filas As Integer
                
                If Valor = True Then
                    
                    Fila = lArrPspto.Count(1) - 1
                    Filas = lArrPspto.Count(1)
                    
                    lArrPspto.ReDim 0, Filas, 0, 18
                    
                    If Fila < 0 Then Fila = 0
                    'On Error Resume Next
                    If CE(lArrPspto(Fila, 0)) <> "" Then
                        Fila = Fila + 1
                    End If
                    
                    If Fila > Filas Then
                        Filas = Filas + 1
                    End If
                    
                    lArrPspto.ReDim 0, Filas, 0, 18
                    
                    lArrPspto(Fila, 0) = CE(param0)
                    lArrPspto(Fila, 1) = CE(param1)
                    '--------------------------------
                    If Right(param0, 4) = "0000" Then
                        lArrPspto(Fila, 2) = "S"
                    Else
                        lArrPspto(Fila, 2) = "N"
                    End If
                    '--------------------------------
                    
                    For i = 3 To 16
                        lArrPspto(Fila, i) = 0
                    Next
                    lArrPspto(Fila, 17) = Me.tdbcMoneda.BoundText
                    
                    
                    '--------------------------------
                    If Right(param0, 4) <> "0000" And Right(param0, 2) = "00" Then
                        lArrPspto(Fila, 18) = "S"
                    Else
                        lArrPspto(Fila, 18) = "N"
                    End If
                    '--------------------------------
                    
                    
                    Set grdEgresos.Array = lArrPspto
                    grdEgresos.ReBind
                    Unload frmBuscador
                    
                    
                    'DoEvents
                    
                    'Grabar
                Else
                    Mensajes "Centro de Costo seleccionado, ya esta contenido actualmente", vbInformation
                End If
                
                Call CalculatotalesFoot
    End Select
End Sub

Private Sub cmdEliminaItem_Click()
    Dim i As Integer
    Dim cad1 As String
    Dim cad2 As String
    Dim Valor As Boolean
    
    Valor = True
    If lArrPspto.Count(1) = 1 And grdEgresos.Bookmark = 0 And CE(grdEgresos.Columns(0)) = "" Then
        Exit Sub
    End If
    
    
    If lArrPspto.Count(1) > 0 Then
'    If lArrPspto(grdEgresos.Bookmark, 2) = "S" Or lArrPspto(grdEgresos.Bookmark, 18) = "S" Then
'        cad1 = CE(grdEgresos.Columns(0))
'        For i = 0 To lArrPspto.Count(1) - 1
'            cad2 = CE(lArrPspto(i, 0))
'            If Len(cad2) > Len(cad1) Then
'                If cad1 = Mid(cad2, 1, Len(cad1)) Then
'                    Valor = False
'                    Exit For
'                End If
'            End If
'        Next
'    End If
    
    If Valor = True And Not IsNull(grdEgresos.Bookmark) Then
        lArrPspto.DeleteRows (grdEgresos.Bookmark)
        grdEgresos.ReBind
        pSetFocus grdEgresos
    End If
    
    If lArrPspto.Count(1) = 0 Then
        lArrPspto.Clear
    End If
    
    If Valor = False Then
        Mensajes "Primero elimine los Centros de costos que hacen referencia a este titulo", vbInformation
    End If
    End If
End Sub

Private Function Grabar() As Boolean
    Grabar = False

    If CE(tdbcMoneda.Text) = "" Then
        Mensajes "Seleccione un tipo de moneda antes de grabar", vbInformation
        Exit Function
    End If

    If lArrPspto.Count(1) = 1 And grdEgresos.Bookmark = 0 And CE(grdEgresos.Columns(0)) = "" Then
        Exit Function
    End If

    Dim i As Integer
    Dim j As Integer
    Dim clsMante As clsMantoTablas
    Screen.MousePointer = vbHourglass
    
'    If gsByMoneda = 1 Then
'        If BuscaTiposCambioPromedio = False And gsByMoneda = 1 Then
'            Screen.MousePointer = vbNormal
'            Exit Function
'        End If
'    End If
    
    'Si no obtiene el Tipo de Cambio Promedio para el Presupuesto
    If gintBiMoneda = 1 Then
        If BuscaTiposCambioPromedio = False And gintBiMoneda = 1 Then
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    End If
    
    ' ***
    Set clsMante = New clsMantoTablas
    ' *** Grabar solo los datos q han sido modificados
    grdEgresos.Bookmark = grdEgresos.Bookmark
    
    ' *** Eliminar los datos anteriores
    ReDim lArrDet(13) As Variant
    lArrDet(0) = "ELIMINAR"             ' Empresa
    lArrDet(1) = gsEmpresa              ' Año
    lArrDet(2) = gsAnio
    lArrDet(3) = ""                     ' *** Periodo
    lTipoPres = "I"
    If Me.optEgresos.Value = True Then lTipoPres = "G"
    lArrDet(4) = lTipoPres              ' *** Tipo
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spPr_GrabaMarcoPres", lArrDet(), False) = False Then
        Screen.MousePointer = vbNormal
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Function
    End If
    ' *** Fin Eliminar los datos anteriores
    
    For i = 0 To lArrPspto.Count(1) - 1
        If CE(lArrPspto(i, 0)) <> "" Then
            For j = 1 To 12

                    Call CargaArregloDet(i, j)
                    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spPr_GrabaMarcoPres", lArrDet(), False) = False Then
                        Screen.MousePointer = vbNormal
                        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                        Exit Function
                    End If

            Next
        End If
    Next
    ' *** Finaliza Clase, ya q aun esta abierta
    clsMante.CommitTrans
    clsMante.FinalizaClase
    Screen.MousePointer = vbNormal
    
    
    tbrOpciones.Buttons(3).Enabled = True
    tbrOpciones.Buttons(5).Enabled = True
    Me.cmdEliminaItem.Enabled = True
    Me.cmdActualiza.Enabled = True
    DoEvents

    Call GeneraArreglo
    
    
    Grabar = True
End Function

Private Function BuscaTiposCambioPromedio() As Boolean
    Dim arrDatos() As Variant
    Dim rsAddItem As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim strMensaje As String

    BuscaTiposCambioPromedio = False
    
    If tdbcMoneda.BoundText = "" Then
        Mensajes "Seleccione la moneda para el proceso de conversion", vbInformation + vbOKOnly
        BuscaTiposCambioPromedio = False
        Exit Function
    End If

    Dim gtxtSQL As String
    gtxtSQL = "select Tca_cEne, Tca_cFeb, Tca_cMar, Tca_cAbr, Tca_cMay, Tca_cJun, " & _
              "Tca_cJul , Tca_cAgo, Tca_cSet, Tca_cOct, Tca_cNov, Tca_cDic " & _
              "From CNT_TIPO_CAMBIO_MENSUAL " & _
              "where tca_ctipo = '0' and emp_ccodigo='" & gsEmpresa & "' and pan_canio ='" & gsAnio & "' and tca_cmoneda='" & gsMonedaExt & "'"
    
    arrDatos = Array(gtxtSQL)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    strMensaje = ""
    
    If Not rsAddItem Is Nothing And rsAddItem.State = adStateOpen Then
        
        If Not (rsAddItem.BOF And rsAddItem.EOF) Then
            Do While Not rsAddItem.EOF
                TCMensual(0) = NE(rsAddItem!Tca_cEne)
                TCMensual(1) = NE(rsAddItem!Tca_cFeb)
                TCMensual(2) = NE(rsAddItem!Tca_cMar)
                TCMensual(3) = NE(rsAddItem!Tca_cAbr)
                TCMensual(4) = NE(rsAddItem!Tca_cMay)
                TCMensual(5) = NE(rsAddItem!Tca_cJun)
                TCMensual(6) = NE(rsAddItem!Tca_cJul)
                TCMensual(7) = NE(rsAddItem!Tca_cAgo)
                TCMensual(8) = NE(rsAddItem!Tca_cSet)
                TCMensual(9) = NE(rsAddItem!Tca_cOct)
                TCMensual(10) = NE(rsAddItem!Tca_cNov)
                TCMensual(11) = NE(rsAddItem!Tca_cDic)
            
                If NE(rsAddItem!Tca_cEne) = 0 Then strMensaje = strMensaje & "ENERO, "
                If NE(rsAddItem!Tca_cFeb) = 0 Then strMensaje = strMensaje & "FEBRERO, "
                If NE(rsAddItem!Tca_cMar) = 0 Then strMensaje = strMensaje & "MARZO, "
                If NE(rsAddItem!Tca_cAbr) = 0 Then strMensaje = strMensaje & "ABRIL, "
                If NE(rsAddItem!Tca_cMay) = 0 Then strMensaje = strMensaje & "MAYO, "
                If NE(rsAddItem!Tca_cJun) = 0 Then strMensaje = strMensaje & "JUNIO, "
                If NE(rsAddItem!Tca_cJul) = 0 Then strMensaje = strMensaje & "JULIO, "
                If NE(rsAddItem!Tca_cAgo) = 0 Then strMensaje = strMensaje & "AGOSTO, "
                If NE(rsAddItem!Tca_cSet) = 0 Then strMensaje = strMensaje & "SETIEMBRE, "
                If NE(rsAddItem!Tca_cOct) = 0 Then strMensaje = strMensaje & "OCTUBRE, "
                If NE(rsAddItem!Tca_cNov) = 0 Then strMensaje = strMensaje & "NOVIEMBRE, "
                If NE(rsAddItem!Tca_cDic) = 0 Then strMensaje = strMensaje & "DICIEMBRE, "
                
                rsAddItem.MoveNext
            Loop
        End If
    Else
        strMensaje = "No hay tipos de cambio mensuales, ingeselos"
        Mensajes strMensaje, vbInformation + vbOKOnly
        Call CerrarRecordSet(rsAddItem)
        Set clDatos = Nothing
        BuscaTiposCambioPromedio = False
        Exit Function
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    
    If strMensaje <> "" Then
        strMensaje = "Faltan ingresar los tipos de cambio promedio de los siguientes meses: " & Chr(10) + Chr(13) & Chr(10) + Chr(13) & strMensaje
        strMensaje = Left(strMensaje, Len(strMensaje) - 2)
        Mensajes strMensaje, vbInformation + vbOKOnly
        
        BuscaTiposCambioPromedio = False
    Else
        BuscaTiposCambioPromedio = True
    End If
    
    
End Function

Private Sub CargaArregloDet(item As Integer, Mes As Integer)
    Dim TC As Double
    
    If gsByMoneda = 1 Then
        TC = TCMensual(Mes - 1)
    Else
        TC = 0
    End If
    
    ReDim lArrDet(13) As Variant
    lArrDet(0) = "INSERTAR"             ' Empresa
    lArrDet(1) = gsEmpresa              ' Año
    lArrDet(2) = gsAnio
    lArrDet(3) = Format(Mes, "00")      ' *** Periodo
    lTipoPres = "I"
    If Me.optEgresos.Value = True Then lTipoPres = "G"
    lArrDet(4) = lTipoPres              ' *** Tipo
    lArrDet(5) = lArrPspto(item, 0)  ' *** Codigo
    lArrDet(6) = tdbcMoneda.BoundText  ' *** Moneda
    
'    If gsByMoneda = 1 Then
    
    If gintBiMoneda = 1 Then
        If Me.tdbcMoneda.Columns(2).Value = 1 Then
            lArrDet(7) = Redondear(lArrPspto(item, Mes + 3), 2)    ' *** MontoSoles
            lArrDet(8) = TC   ' *** TC
            lArrDet(9) = Redondear(lArrPspto(item, Mes + 3) / IIf(TC = 0, 1, TC), 2) ' *** MontoDolares
        Else
            lArrDet(7) = Redondear(lArrPspto(item, Mes + 3) * IIf(TC = 0, 1, TC), 2)  ' *** MontoSoles
            lArrDet(8) = TC ' *** TC
            lArrDet(9) = lArrPspto(item, Mes + 3) ' *** MontoDolares
        End If
    Else

        lArrDet(7) = Redondear(lArrPspto(item, Mes + 3), 2)     ' *** MontoSoles
        lArrDet(8) = 0   ' *** TC
        lArrDet(9) = 0 ' *** MontoDolares

    End If
    'lArrDet(10) = "" ' *** Fecha
    lArrDet(11) = "" ' *** Observacion
    lArrDet(12) = "A"                   ' *** Estado
    lArrDet(13) = gsUsuario             ' *** Usuario
End Sub

Private Sub Imprimir()
    frmRepPresupuestoListado.Show
End Sub

Private Sub cmdListar_Click()
    Call GeneraArreglo  ' *** Cargar datos de Ingresos
End Sub

Private Sub cmdOrdenar_Click()
    Grabar
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Dim sqlcombos As String
    grdEgresos.FetchRowStyle = True
    
    Centrar_form Me
    
    lblanio = "AÑO: " & gsAnio
    ' *** Llenando el tipo de Moneda
    Dim strMon As String
'    If gsByMoneda = 1 Then
'        strMon = " (Mon_cMNac = '1' or Mon_cMExt = '1') "
'    Else
'        strMon = " Mon_cMNac = '1' "
'    End If
    
    'Muestra en el listado las moneda Soles y Dolares
    If gintBiMoneda = 1 Then
        strMon = " (Mon_cMNac = '1' or Mon_cMExt = '1') "
    Else
        strMon = " Mon_cMNac = '1' "
    End If
    
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND " & strMon & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    ' ***
    Call GeneraArreglo  ' *** Generar Arreglo
    
    tbrOpciones.Buttons(3).Enabled = True
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    
    'grdEgresos.Splits(0).Locked = False
    grdEgresos.Splits(0).MarqueeStyle = dbgHighlightRow
    grdEgresos.HighlightRowStyle = "HighlightRow"
    
    
    grdEgresos.ReBind
    
    
    tdbcMoneda.Locked = False
    
'    tdbnTipoCambio.ReadOnly = False

    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdEliminaItem.Enabled = False
        cmdActualiza.Enabled = False
        cmdOrdenar.Enabled = False
        tdbcMoneda.Enabled = False
        
        grdEgresos.Splits(1).Locked = True
    Else
        cmdEliminaItem.Enabled = True
        cmdActualiza.Enabled = True
        cmdOrdenar.Enabled = True
        tdbcMoneda.Enabled = True
        
        grdEgresos.Splits(1).Locked = False
    End If



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If tbrOpciones.Buttons(3).Enabled = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                'If respuesta = vbYes Then Call Cancelar
            End If
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        'Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub

Private Sub GeneraArreglo()
    ' ***
    Dim sqlPres As String
    Dim i As Integer
    
    lTipoPres = "I"
    If Me.optEgresos.Value = True Then lTipoPres = "G"
    ' *** Llenar Datos
    On Local Error GoTo ErrorEjecucion
    sqlPres = "spPr_ConsultaPresupuestos 'SEL_ALLTIPO', '" & gsEmpresa & "', '" & gsAnio & _
                "', '', '" & lTipoPres & "', '', ''"
    
    Call GridArreglo(lArrPspto, Me.grdEgresos, sqlPres)
    
    
    If lArrPspto.UpperBound(1) > 0 Then
        'Me.tdbnTipoCambio.Value = lArrPspto(0, 16)
        'If lArrPspto(0, 17) <> "" Then
        tdbcMoneda.BoundText = lArrPspto(0, 17)
    End If
    
    Call CalculatotalesFoot

    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " : " & Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        grdEgresos.Height = Me.Height - 2600
        grdEgresos.Width = Me.Width - 300
        tbrOpciones.Width = Me.Width
        Exit Sub
    End If
    Exit Sub
serror:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lArrPspto = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

Private Sub grdEgresos_AfterColEdit(ByVal ColIndex As Integer)
    Dim Fila As Integer
    If ColIndex >= 3 Then
        grdEgresos.Columns(ColIndex) = NE(grdEgresos.Columns(ColIndex))
        
        If grdEgresos.Columns(ColIndex) = "" Then grdEgresos.Columns(ColIndex) = 0
        lArrPspto(grdEgresos.Bookmark, 3) = (grdEgresos.Columns(3).Value - NE(valorPres)) + grdEgresos.Columns(ColIndex).Value
        grdEgresos.Columns(3) = lArrPspto(grdEgresos.Bookmark, 3)
        grdEgresos.Update

    End If


    Call SumarTotales(3)
    Call SumarTotales(ColIndex)
End Sub


Private Sub CalculatotalesFoot()
    Dim Col As Integer
    If lArrPspto.Count(1) > 1 Then
    
        For Col = 3 To 15
             SumarTotales (Col)
        Next
    
    End If
    grdEgresos.Refresh
End Sub
Private Sub grdEgresos_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

    
    If ColIndex = 3 Then
        Cancel = 1
        Exit Sub
    End If
    
    If ColIndex >= 3 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
            Cancel = 0
        Else
            Cancel = 1
            Exit Sub
        End If
    End If
    
    If ColIndex < 3 Then
        Cancel = 1
        Exit Sub
    End If
    
    If ColIndex > 2 Then
        If grdEgresos.Columns(ColIndex).Value = "" Then grdEgresos.Columns(ColIndex).Value = 0
        valorPres = grdEgresos.Columns(ColIndex).Value
    End If
End Sub

Private Sub DivideMontos(j As Integer)
    Dim k As Integer
    
    For k = 4 To 15
        grdEgresos.Columns(k) = grdEgresos.Columns(j) / 12
    Next
End Sub

Private Sub DistribuyePartida(Valor As Integer)
    Dim k As Integer
    For k = 4 To 15
        lArrPspto(Valor, k) = lArrPspto(Valor, 3) / 12
    Next
End Sub

Private Sub optEgresos_Click()
    Call GeneraArreglo  ' *** Cargar datos de Egresos
End Sub

Private Sub optIngresos_Click()
    Call GeneraArreglo  ' *** Cargar datos de Ingresos
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 3:
                If Grabar = True Then
                   On Error Resume Next
                   grdEgresos.Update
                   DoEvents
                   Mensajes "El presupuesto se grabo con exito", vbInformation
                   
                   grdEgresos.Refresh
                   
                End If
                
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                
        Case 5: 'Editar
        Case 6: Imprimir
        Case 7
            'If tbrOpciones.Buttons(3).Enabled = False Then ' *** Grabar
                Unload Me
'            Else
'                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
'                If respuesta = vbYes Then
'                    Call Cancelar
'                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
'                End If
'            End If
    End Select
End Sub

Private Sub SumarTotales(Columna As Integer)
    Dim s_suma As Double
    Dim i As Integer
    If Columna > 1 Then
        s_suma = 0
        grdEgresos.Update
        For i = 0 To lArrPspto.Count(1) - 1
            'If lArrPspto(i, 2) = "N" And lArrPspto(i, 18) = "N" Then
                s_suma = s_suma + Redondear(lArrPspto(i, Columna), 2)
            'End If
        Next
        grdEgresos.Columns(Columna).FooterText = Format(s_suma, "###,###,##0.00")
    End If
End Sub

