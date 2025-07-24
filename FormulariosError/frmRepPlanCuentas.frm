VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepPlanCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Plan de Cuentas"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   Icon            =   "frmRepPlanCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   7440
   Begin VB.Frame fraTodo 
      Height          =   4200
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7350
      Begin VB.Frame Frame2 
         Height          =   2520
         Left            =   3630
         TabIndex        =   14
         Top             =   180
         Width           =   3555
         Begin VB.OptionButton optTodo 
            Caption         =   "Todo Plan de Cuentas"
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
            Left            =   450
            TabIndex        =   17
            Top             =   675
            Value           =   -1  'True
            Width           =   2595
         End
         Begin VB.OptionButton optRango 
            Caption         =   "Por Rango de Cuentas"
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
            Left            =   450
            TabIndex        =   16
            Top             =   1260
            Width           =   2595
         End
         Begin VB.Frame Frame4 
            Height          =   105
            Index           =   0
            Left            =   405
            TabIndex        =   15
            Top             =   1035
            Width           =   2355
         End
         Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
            Height          =   315
            Left            =   1305
            TabIndex        =   18
            Tag             =   "_"
            Top             =   1650
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "frmRepPlanCuentas.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepPlanCuentas.frx":0F36
            Key             =   "frmRepPlanCuentas.frx":0F54
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
         Begin TDBText6Ctl.TDBText tdbtCuentaHasta 
            Height          =   315
            Left            =   1305
            TabIndex        =   19
            Tag             =   "_"
            Top             =   2010
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "frmRepPlanCuentas.frx":0FA6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepPlanCuentas.frx":1012
            Key             =   "frmRepPlanCuentas.frx":1030
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   225
            Left            =   480
            TabIndex        =   22
            Top             =   2055
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Height          =   225
            Left            =   480
            TabIndex        =   21
            Top             =   1695
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RANGO DE CUENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   900
            TabIndex        =   20
            Top             =   180
            Width           =   1605
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1185
         Left            =   180
         TabIndex        =   10
         Top             =   2880
         Width           =   3420
         Begin VB.CheckBox chkDigitosHasta 
            Caption         =   "Mostrar hasta los digitos ingresados"
            Height          =   285
            Left            =   135
            TabIndex        =   11
            Top             =   765
            Width           =   3120
         End
         Begin TDBText6Ctl.TDBText tdbnDigitos 
            Height          =   315
            Left            =   2655
            TabIndex        =   12
            Tag             =   "_"
            Top             =   270
            Width           =   570
            _Version        =   65536
            _ExtentX        =   1005
            _ExtentY        =   556
            Caption         =   "frmRepPlanCuentas.frx":1082
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepPlanCuentas.frx":10EE
            Key             =   "frmRepPlanCuentas.frx":110C
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
            AlignHorizontal =   2
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
            Text            =   "2"
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
            Caption         =   "Indique el numero de digitos a mostrar en el reporte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   135
            TabIndex        =   13
            Top             =   225
            Width           =   2490
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2685
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   3420
         Begin VB.OptionButton optConasev 
            Caption         =   "Impresión de Cuentas"
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
            Left            =   180
            TabIndex        =   6
            Top             =   480
            Width           =   2235
         End
         Begin VB.OptionButton optDestinoResumen 
            Caption         =   "Impresión con Destino"
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
            Left            =   180
            TabIndex        =   5
            Top             =   1935
            Value           =   -1  'True
            Width           =   2235
         End
         Begin VB.OptionButton optDestinoDetalle 
            Caption         =   "Impresión con Destino Anual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   4
            Top             =   810
            Width           =   2910
         End
         Begin VB.Frame Frame4 
            Height          =   105
            Index           =   1
            Left            =   165
            TabIndex        =   3
            Top             =   1725
            Width           =   2985
         End
         Begin VB.OptionButton optDestinoSinAmarre 
            Caption         =   "Cuentas con Destino sin Amarre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   180
            TabIndex        =   2
            Top             =   1050
            Width           =   3120
         End
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   960
            TabIndex        =   7
            Top             =   2250
            Width           =   2295
            _ExtentX        =   4048
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
            _PropDict       =   $"frmRepPlanCuentas.frx":115E
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TIPO DE REPORTE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1035
            TabIndex        =   9
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
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
            Left            =   465
            TabIndex        =   8
            Top             =   2295
            Width           =   345
         End
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   3720
         TabIndex        =   24
         Top             =   3540
         Width           =   1665
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepPlanCuentas.frx":11E5
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   5475
         TabIndex        =   23
         Top             =   3540
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepPlanCuentas.frx":177F
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
      TabIndex        =   25
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Control As String
Private Sub cmdImprimir_Click()
    ' *** Imprime la plantilla del Balance
   'Validar Ingreso ''
    cmdImprimir.Enabled = True
    Screen.MousePointer = vbHourglass
    DoEvents
    If optRango.Value = True Then
        Dim gtxtSQL As String
        Dim obj As ClsFuncionesExecute
        Set obj = New ClsFuncionesExecute
        
        If CE(tdbtCuentaDesde.Text) = "" Or Trim(tdbtCuentaDesde.Text) = "" Then
            Mensajes "Verificar : Faltan datos   "
            Set obj = Nothing
            cmdImprimir.Enabled = True
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        If NE(tdbtCuentaDesde.Text) > NE(tdbtCuentaHasta.Text) Then
            Mensajes "Verificar el Valor de la cuenta inicio no puede ser mayor a la cuenta final"
            Set obj = Nothing
            cmdImprimir.Enabled = True
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        gtxtSQL = "SELECT  * FROM CNM_PLAN_CTA WHERE Emp_cCodigo ='" & gsEmpresa & "' AND Pan_cAnio ='" & gsAnio & "' AND Pla_cCuentaContable='" & tdbtCuentaDesde & "'"
        If obj.fRetornaRS(gtxtSQL).RecordCount = 0 Then
            Mensajes "Verificar : No existe la cuenta Desde"
            Set obj = Nothing
            cmdImprimir.Enabled = True
            Screen.MousePointer = vbNormal
            pSetFocus tdbtCuentaDesde
            Exit Sub
        End If
            
        gtxtSQL = "SELECT  * FROM CNM_PLAN_CTA WHERE Emp_cCodigo ='" & gsEmpresa & "' AND Pan_cAnio ='" & gsAnio & "' AND Pla_cCuentaContable='" & tdbtCuentaHasta & "'"
        If obj.fRetornaRS(gtxtSQL).RecordCount = 0 Then
            Mensajes "Verificar : No existe la cuenta Hasta"
            Set obj = Nothing
            cmdImprimir.Enabled = True
            Screen.MousePointer = vbNormal
            pSetFocus tdbtCuentaHasta
            Exit Sub
        End If
    
    End If
     
    Dim matriz_fecha(9) As Variant
    If Me.optConasev.Value = True Then
        If chkDigitosHasta.Value = vbChecked Then
            matriz_fecha(0) = "@Tipo;SEL_ALL_HASTA;True"
        Else
            matriz_fecha(0) = "@Tipo;SEL_ALL;True"
        End If
    Else
        matriz_fecha(0) = "@Tipo;RPTSEL_ALLDESTDET;True"
    End If
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Pla_cAnioPlan;" & gsAnio & ";True"
    
    If optTodo.Value = True Then
        matriz_fecha(3) = "@Pla_cCuentaContable1;00;True"
        matriz_fecha(4) = "@Pla_cCuentaContable2;999999999999;True"
    Else
        matriz_fecha(3) = "@Pla_cCuentaContable1;" & Me.tdbtCuentaDesde & ";True"
        matriz_fecha(4) = "@Pla_cCuentaContable2;" & Me.tdbtCuentaHasta & ";True"
    End If
    
    matriz_fecha(5) = "@Periodo;" & Me.tdbcMes.BoundText & ";True"
    matriz_fecha(6) = "@Digitos;" & Me.tdbnDigitos.Text & ";True"
    
    matriz_fecha(7) = "@DigHasta;" & Me.tdbnDigitos.Text & ";True"
    
    matriz_fecha(8) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(9) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    
    If Me.optDestinoResumen.Value = True Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentasContables.rpt", crptToWindow, "Plan de Cuentas", "", matriz_fecha(), formulas()
    Else
        If Me.optDestinoDetalle.Value = True Then
            matriz_fecha(0) = "@Tipo;RPTSEL_ALLDESTANU;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentasContablesDetAnual.rpt", crptToWindow, "Plan de Cuentas", "", matriz_fecha(), formulas()
        
        ElseIf Me.optDestinoSinAmarre.Value = True Then
            matriz_fecha(0) = "@Tipo;RPTSEL_SINDESTANU;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentasContablesDetAnualSinDest.rpt", crptToWindow, "Plan de Cuentas", "", matriz_fecha(), formulas()
        Else
            AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentasContablesConasev.rpt", crptToWindow, "Plan de Cuentas", "", matriz_fecha(), formulas()
        End If
    End If
    
    Set obj = Nothing
    cmdImprimir.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Centrar_form(Me)
    
    
    tdbtCuentaDesde.ReadOnly = True
    tdbtCuentaHasta.ReadOnly = True
    tdbtCuentaDesde.BackColor = gsColorDesactivado
    tdbtCuentaHasta.BackColor = gsColorDesactivado
    
    Call LlenaComboMesAddItem(tdbcMes)
    tdbcMes.BoundText = gsPeriodo
    optConasev.Value = vbChecked
    On Error Resume Next
    tdbcMes.ReBind
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

Private Sub optConasev_Click()
    Me.tdbcMes.Enabled = False
    tdbnDigitos.Enabled = True
    chkDigitosHasta.Enabled = True
End Sub

Private Sub optDestinoDetalle_Click()
    tdbnDigitos.Text = "8"
    tdbnDigitos.Enabled = False
    chkDigitosHasta.Enabled = False
End Sub

Private Sub optDestinoResumen_Click()
    Me.tdbcMes.Enabled = True
    tdbnDigitos.Enabled = True
End Sub

Private Sub optDestinoSinAmarre_Click()
    tdbnDigitos.Text = "8"
    tdbnDigitos.Enabled = False
    chkDigitosHasta.Enabled = False
End Sub

Private Sub optRango_Click()

tdbtCuentaDesde.ReadOnly = False
tdbtCuentaHasta.ReadOnly = False
tdbtCuentaDesde.Enabled = True
tdbtCuentaHasta.Enabled = True
tdbtCuentaDesde.BackColor = gsColorActivado
tdbtCuentaHasta.BackColor = gsColorActivado
pSetFocus tdbtCuentaDesde

End Sub
Private Sub optTodo_Click()

tdbtCuentaDesde.Text = ""
tdbtCuentaHasta.Text = ""

tdbtCuentaDesde.ReadOnly = True
tdbtCuentaHasta.ReadOnly = True
tdbtCuentaDesde.Enabled = False
tdbtCuentaHasta.Enabled = False
tdbtCuentaDesde.BackColor = gsColorDesactivado
tdbtCuentaHasta.BackColor = gsColorDesactivado

End Sub


Private Sub tdbnDigitos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If ValidaDigitos(tdbnDigitos.Text) = False Then
            tdbnDigitos.Text = "2"
            pSetFocus tdbnDigitos
        End If
    End If
End Sub

Private Sub tdbnDigitos_LostFocus()
    If ValidaDigitos(tdbnDigitos.Text) = False Then
        tdbnDigitos.Text = "2"
        pSetFocus tdbnDigitos
    End If
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And tdbtCuentaDesde.ReadOnly = False Then
    Call LlamaBuscar(frmBuscador, Me.tdbtCuentaDesde.Name, Control, "CuentasN", Me, tdbcMes.BoundText, tdbtCuentaDesde)
End If
End Sub
Private Sub tdbtCuentaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And tdbtCuentaHasta.ReadOnly = False Then
    Call LlamaBuscar(frmBuscador, Me.tdbtCuentaHasta.Name, Control, "CuentasN", Me, tdbcMes.BoundText, tdbtCuentaHasta)
End If
End Sub
Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
            Case "tdbtCuentaDesde"   ' *** Caso de cliente
                 tdbtCuentaDesde = Trim(param0)
                 'Me.tdbtNombreDestino = Trim(frmBuscador.TDBGTabla.Columns(1).Value)
                 Unload frmBuscador
                 pSetFocus tdbtCuentaDesde
             Case Else: ' Case "tdbtCuentahasta"
                 tdbtCuentaHasta = Trim(param0)
                 Unload frmBuscador
                 pSetFocus tdbtCuentaHasta
    End Select
End Sub

