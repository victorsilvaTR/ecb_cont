VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Empresas"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "frmManEmpresas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   7620
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   0
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
            Picture         =   "frmManEmpresas.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":220C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":25E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":29C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5820
      Left            =   90
      TabIndex        =   29
      Top             =   450
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   10266
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Empresas"
      TabPicture(0)   =   "frmManEmpresas.frx":39DA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Empresa"
      TabPicture(1)   =   "frmManEmpresas.frx":39F6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblMante"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Configuración"
      TabPicture(2)   =   "frmManEmpresas.frx":3A12
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4(2)"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "tdbcEmpresa"
      Tab(2).Control(3)=   "tdbcAnio"
      Tab(2).Control(4)=   "tdbcEmpresaDest"
      Tab(2).Control(5)=   "tdbcAnioDest"
      Tab(2).Control(6)=   "cmdEliminarEjercicio"
      Tab(2).Control(7)=   "cmdRefreshDestino"
      Tab(2).Control(8)=   "cmdRefresh"
      Tab(2).Control(9)=   "cmdExportar"
      Tab(2).Control(10)=   "Label2(15)"
      Tab(2).Control(11)=   "Label2(14)"
      Tab(2).ControlCount=   12
      Begin VB.Frame Frame4 
         Height          =   1905
         Index           =   2
         Left            =   -69690
         TabIndex        =   61
         Top             =   3825
         Width           =   60
      End
      Begin VB.Frame Frame3 
         Caption         =   " CONFIGURACIÓN INICIAL "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   -74865
         TabIndex        =   46
         Top             =   450
         Width           =   7215
         Begin VB.CheckBox chkDetalleEnt 
            Caption         =   "Detalle de Entidades"
            Height          =   285
            Left            =   2385
            TabIndex        =   65
            Top             =   1935
            Width           =   1905
         End
         Begin VB.CheckBox chkCenCos 
            Caption         =   "Centro de costo"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4995
            TabIndex        =   24
            Top             =   1620
            Width           =   2040
         End
         Begin VB.CheckBox chkPlan 
            Caption         =   "Plan de cuentas"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   180
            TabIndex        =   10
            Top             =   720
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkLibros 
            Caption         =   "Libros"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   180
            TabIndex        =   12
            Top             =   1305
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkEntidades 
            Caption         =   "Entidades"
            Height          =   285
            Left            =   180
            TabIndex        =   13
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTipDoc 
            Caption         =   "Tipo de documentos"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTabSec 
            Caption         =   "Tablas secundarias"
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTipMon 
            Caption         =   "Tipo de moneda"
            Height          =   285
            Left            =   180
            TabIndex        =   16
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.Frame Frame4 
            Height          =   2895
            Index           =   0
            Left            =   2115
            TabIndex        =   51
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkTipAsto 
            Caption         =   "Tipo de asiento"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2385
            TabIndex        =   20
            Top             =   1305
            Width           =   1905
         End
         Begin VB.CheckBox chkPlanCONA 
            Caption         =   "Plantilla EEFF"
            Height          =   285
            Left            =   180
            TabIndex        =   15
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkParamIni 
            Caption         =   "Parametros iniciales"
            Height          =   285
            Left            =   2385
            TabIndex        =   18
            Top             =   675
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkConfigOp 
            Caption         =   "Config. de operaciones"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2385
            TabIndex        =   19
            Top             =   990
            Value           =   1  'Checked
            Width           =   1995
         End
         Begin VB.Frame Frame4 
            Height          =   2895
            Index           =   1
            Left            =   4680
            TabIndex        =   47
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkBancos 
            Caption         =   "Bancos"
            Height          =   285
            Left            =   4995
            TabIndex        =   21
            Top             =   720
            Width           =   1905
         End
         Begin VB.CheckBox chkCtaCte 
            Caption         =   "Cuenta corriente"
            Height          =   285
            Left            =   4995
            TabIndex        =   22
            Top             =   990
            Width           =   1905
         End
         Begin VB.CheckBox chkRatios 
            Caption         =   "Ratios financieros"
            Height          =   285
            Left            =   4995
            TabIndex        =   23
            Top             =   1305
            Width           =   2040
         End
         Begin VB.CheckBox chkTipoCbio 
            Caption         =   "Tipo de cambio"
            Height          =   285
            Left            =   4995
            TabIndex        =   25
            Top             =   1935
            Width           =   1500
         End
         Begin VB.Frame frmFechas 
            Height          =   1050
            Left            =   4860
            TabIndex        =   48
            Top             =   1980
            Width           =   2220
            Begin TDBDate6Ctl.TDBDate dtpFechaIni 
               Height          =   300
               Left            =   720
               TabIndex        =   26
               Tag             =   "enabled"
               Top             =   270
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   529
               Calendar        =   "frmManEmpresas.frx":3A2E
               Caption         =   "frmManEmpresas.frx":3B30
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEmpresas.frx":3B94
               Keys            =   "frmManEmpresas.frx":3BB2
               Spin            =   "frmManEmpresas.frx":3C1E
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
               Enabled         =   0
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
            Begin TDBDate6Ctl.TDBDate dtpFechaFin 
               Height          =   300
               Left            =   720
               TabIndex        =   27
               Tag             =   "enabled"
               Top             =   630
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   529
               Calendar        =   "frmManEmpresas.frx":3C46
               Caption         =   "frmManEmpresas.frx":3D48
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManEmpresas.frx":3DAC
               Keys            =   "frmManEmpresas.frx":3DCA
               Spin            =   "frmManEmpresas.frx":3E36
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
               Enabled         =   0
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
            Begin VB.Label Label2 
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
               Index           =   11
               Left            =   90
               TabIndex        =   50
               Top             =   270
               Width           =   555
            End
            Begin VB.Label Label2 
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
               Index           =   12
               Left            =   90
               TabIndex        =   49
               Top             =   630
               Width           =   495
            End
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnDigitos 
            Height          =   300
            Left            =   1080
            TabIndex        =   11
            Top             =   990
            Width           =   540
            _Version        =   65536
            _ExtentX        =   952
            _ExtentY        =   529
            Calculator      =   "frmManEmpresas.frx":3E5E
            Caption         =   "frmManEmpresas.frx":3E7E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":3EE2
            Keys            =   "frmManEmpresas.frx":3F00
            Spin            =   "frmManEmpresas.frx":3F58
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   12
            MinValue        =   2
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   3
            MaxValueVT      =   1802698757
            MinValueVT      =   1769209861
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Digitos"
            Height          =   195
            Index           =   16
            Left            =   450
            TabIndex        =   55
            Top             =   1035
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PRINCIPALES"
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
            Index           =   9
            Left            =   180
            TabIndex        =   54
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "DEPENDIENTES"
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
            Index           =   10
            Left            =   2385
            TabIndex        =   53
            Top             =   405
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "OPCIONALES"
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
            Index           =   13
            Left            =   4995
            TabIndex        =   52
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5265
         Left            =   -74865
         TabIndex        =   36
         Top             =   405
         Width           =   7230
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   3945
            Left            =   405
            TabIndex        =   1
            Top             =   1080
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   6959
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo "
            Columns(0).DataField=   "Emp_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción "
            Columns(1).DataField=   "Emp_cNombreLargo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Sucursal"
            Columns(2).DataField=   "Emp_cCodSuc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6429"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6350"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Named:id=33:Normal"
            _StyleDefs(51)  =   ":id=33,.parent=0"
            _StyleDefs(52)  =   "Named:id=34:Heading"
            _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(54)  =   ":id=34,.wraptext=-1"
            _StyleDefs(55)  =   "Named:id=35:Footing"
            _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=36:Selected"
            _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=37:Caption"
            _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(61)  =   "Named:id=38:HighlightRow"
            _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(63)  =   "Named:id=39:EvenRow"
            _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(65)  =   "Named:id=40:OddRow"
            _StyleDefs(66)  =   ":id=40,.parent=33"
            _StyleDefs(67)  =   "Named:id=41:RecordSelector"
            _StyleDefs(68)  =   ":id=41,.parent=34"
            _StyleDefs(69)  =   "Named:id=42:FilterBar"
            _StyleDefs(70)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1980
            TabIndex        =   0
            Top             =   675
            Width           =   4110
            _Version        =   65536
            _ExtentX        =   7250
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":3F80
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":3FEC
            Key             =   "frmManEmpresas.frx":400A
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
            Caption         =   "Nombre Empresa"
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
            Left            =   360
            TabIndex        =   38
            Top             =   720
            Width           =   1485
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
            Left            =   360
            TabIndex        =   37
            Top             =   315
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5190
         Left            =   180
         TabIndex        =   30
         Top             =   450
         Width           =   7110
         Begin VB.CheckBox chkBiMoneda 
            Caption         =   "BiMoneda"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Top             =   3995
            Width           =   1215
         End
         Begin VB.CommandButton cmdEmpresas 
            Height          =   330
            Left            =   2835
            Picture         =   "frmManEmpresas.frx":405C
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   " Seleccione una crempresa"
            Top             =   1080
            Width           =   375
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   1800
            TabIndex        =   2
            Tag             =   "_"
            Top             =   1080
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":45E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":4652
            Key             =   "frmManEmpresas.frx":4670
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
         Begin TDBText6Ctl.TDBText tdbtNombreCorto 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Tag             =   "_"
            Top             =   1890
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":46C2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":472E
            Key             =   "frmManEmpresas.frx":474C
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
            MaxLength       =   20
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
         Begin TDBText6Ctl.TDBText tdbtDireccion 
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Tag             =   "_"
            Top             =   2295
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":479E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":480A
            Key             =   "frmManEmpresas.frx":4828
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
            MaxLength       =   80
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
            Left            =   1800
            TabIndex        =   7
            Tag             =   "_"
            Top             =   2700
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":487A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":48E6
            Key             =   "frmManEmpresas.frx":4904
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
            MaxLength       =   11
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
         Begin TDBText6Ctl.TDBText tdbtTelefono 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Tag             =   "_"
            Top             =   3105
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":4948
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":49B4
            Key             =   "frmManEmpresas.frx":49D2
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   30
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
         Begin TrueOleDBList70.TDBCombo tdbcSucursal 
            Height          =   300
            Left            =   1800
            TabIndex        =   9
            Tag             =   "_"
            Top             =   3510
            Width           =   3930
            _ExtentX        =   6932
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
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
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
            _PropDict       =   $"frmManEmpresas.frx":4A24
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
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtNombreLargo 
            Height          =   315
            Left            =   1800
            TabIndex        =   4
            Tag             =   "_"
            Top             =   1485
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   556
            Caption         =   "frmManEmpresas.frx":4AAB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":4B17
            Key             =   "frmManEmpresas.frx":4B35
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
            MaxLength       =   80
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
         Begin TDBNumber6Ctl.TDBNumber tdbtAnio 
            Height          =   300
            Left            =   4770
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   529
            Calculator      =   "frmManEmpresas.frx":4B87
            Caption         =   "frmManEmpresas.frx":4BA7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManEmpresas.frx":4C0B
            Keys            =   "frmManEmpresas.frx":4C29
            Spin            =   "frmManEmpresas.frx":4C81
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   2100
            MinValue        =   1900
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   2004
            MaxValueVT      =   1802698757
            MinValueVT      =   1769209861
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E M P R E S A"
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
            Index           =   17
            Left            =   2700
            TabIndex        =   56
            Top             =   495
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal"
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
            Index           =   7
            Left            =   285
            TabIndex        =   44
            Top             =   3555
            Width           =   735
         End
         Begin VB.Label lblanio 
            AutoSize        =   -1  'True
            Caption         =   "Año de Inicio"
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
            Left            =   3465
            TabIndex        =   43
            Top             =   1125
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
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
            Left            =   285
            TabIndex        =   42
            Top             =   3150
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ruc"
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
            Left            =   285
            TabIndex        =   41
            Top             =   2745
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
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
            Left            =   285
            TabIndex        =   40
            Top             =   2340
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Corto"
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
            Left            =   285
            TabIndex        =   39
            Top             =   1980
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Largo"
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
            Left            =   285
            TabIndex        =   35
            Top             =   1530
            Width           =   1200
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
            Left            =   285
            TabIndex        =   34
            Top             =   1125
            Width           =   600
         End
      End
      Begin TrueOleDBList70.TDBCombo tdbcEmpresa 
         Height          =   300
         Left            =   -74865
         TabIndex        =   57
         Tag             =   "enabled"
         Top             =   4185
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   9155
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
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=847"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=767"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
         CellTips        =   1
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
         _PropDict       =   $"frmManEmpresas.frx":4CA9
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
      Begin TrueOleDBList70.TDBCombo tdbcAnio 
         Height          =   300
         Left            =   -71805
         TabIndex        =   28
         Tag             =   "enabled"
         Top             =   4185
         Width           =   1485
         _ExtentX        =   2619
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
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=159"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=79"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1138"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1058"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   4
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
         _PropDict       =   $"frmManEmpresas.frx":4D30
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
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBList70.TDBCombo tdbcEmpresaDest 
         Height          =   300
         Left            =   -74865
         TabIndex        =   59
         Tag             =   "enabled"
         Top             =   5040
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   9155
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
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=847"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=767"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
         CellTips        =   1
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
         _PropDict       =   $"frmManEmpresas.frx":4DB7
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
      Begin TrueOleDBList70.TDBCombo tdbcAnioDest 
         Height          =   300
         Left            =   -71805
         TabIndex        =   31
         Tag             =   "enabled"
         Top             =   5040
         Width           =   1485
         _ExtentX        =   2619
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
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=159"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=79"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1138"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1058"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   4
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
         _PropDict       =   $"frmManEmpresas.frx":4E3E
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
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSForms.CommandButton cmdEliminarEjercicio 
         Height          =   435
         Left            =   -69330
         TabIndex        =   64
         Top             =   4815
         Visible         =   0   'False
         Width           =   1665
         Caption         =   " Elimina Año"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefreshDestino 
         Height          =   390
         Left            =   -70275
         TabIndex        =   32
         ToolTipText     =   "Cargar Lista"
         Top             =   4950
         Width           =   405
         PicturePosition =   327683
         Size            =   "714;688"
         Picture         =   "frmManEmpresas.frx":4EC5
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   -70275
         TabIndex        =   62
         Top             =   4140
         Visible         =   0   'False
         Width           =   405
         PicturePosition =   327683
         Size            =   "714;688"
         Picture         =   "frmManEmpresas.frx":5B17
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdExportar 
         Height          =   435
         Left            =   -69330
         TabIndex        =   33
         Top             =   4140
         Width           =   1665
         Caption         =   "Importar Datos"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NUEVA EMPRESA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   15
         Left            =   -74865
         TabIndex        =   60
         Top             =   4725
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMPRESA DE ORIGEN"
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
         Index           =   14
         Left            =   -74865
         TabIndex        =   58
         Top             =   3870
         Width           =   1800
      End
      Begin VB.Label lblMante 
         Height          =   285
         Left            =   9900
         TabIndex        =   45
         Top             =   765
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":60B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":620B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":6365
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":64BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":6619
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":6773
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":68CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":6A27
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":6B81
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":711B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":76B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":7C4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":81E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":8783
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":8D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManEmpresas.frx":92B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   66
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir F7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar o Salir ESC"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lArrAnio As New XArrayDB
Dim lsFecha As Date
Dim CodigoEmp As String
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
    '''''''''''''''''''''''''''''''''''''
    Dim CMD As ADODB.Command
    Dim rs_plan_cta As ADODB.Recordset
    ''''''''''''''''''''''''''''''''''''
    
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkPlan_Click()
    If chkPlan.Value = vbChecked Then
        tdbnDigitos.Enabled = True
    Else
        tdbnDigitos.Enabled = False
    End If
End Sub

Private Sub chkTipoCbio_Click()
    If chkTipoCbio.Value = vbChecked Then
        
        ActivarControl dtpFechaIni, True
        ActivarControl dtpFechaFin, True
        
        On Error Resume Next
        dtpFechaIni.Value = PrimerDiaMes("01", tdbcAnio.BoundText)
        dtpFechaFin.Value = UltimoDiaMes("12", tdbcAnio.BoundText)
    
    Else
        ActivarControl dtpFechaIni, False
        ActivarControl dtpFechaFin, False
    
        On Error Resume Next
        dtpFechaIni.Value = "03/08/2004"
        dtpFechaFin.Value = "03/08/2004"
    End If
End Sub

Private Function Validar() As Boolean

    If tdbcEmpresaDest.BoundText = gsEmpresa Then
        Mensajes "La empresa de destino no puede ser la misma empresa a la que inicio sesion", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If


    If tdbcEmpresaDest.BoundText = tdbcEmpresa.BoundText Then
        Mensajes "La empresa de origen y la empresa de destino no pueden ser las mismas", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If


    If tdbcEmpresaDest.Enabled = False Then
        Mensajes "Debe seleccionar como minimo una empresa de destino", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If

    If chkBancos.Value = vbUnchecked And _
       chkConfigOp.Value = vbUnchecked And _
       chkCtaCte.Value = vbUnchecked And _
       chkEntidades.Value = vbUnchecked And _
       chkLibros.Value = vbUnchecked And _
       chkParamIni.Value = vbUnchecked And _
       chkPlan.Value = vbUnchecked And _
       chkPlanCONA.Value = vbUnchecked And _
       chkRatios.Value = vbUnchecked And _
       chkTabSec.Value = vbUnchecked And _
       chkTipAsto.Value = vbUnchecked And _
       chkTipDoc.Value = vbUnchecked And _
       chkTipMon.Value = vbUnchecked And _
       chkCenCos.Value = vbUnchecked And _
       chkTipoCbio.Value = vbUnchecked Then
    
        Mensajes "Seleccione minimo una tabla de la lista", vbOKOnly + vbInformation
        Validar = False
        Exit Function
    
    End If
    

    If chkBancos.Value = vbUnchecked And _
       chkConfigOp.Value = vbUnchecked And _
       chkCtaCte.Value = vbUnchecked And _
       chkEntidades.Value = vbUnchecked And _
       chkLibros.Value = vbUnchecked And _
       chkParamIni.Value = vbUnchecked And _
       chkPlan.Value = vbUnchecked And _
       chkPlanCONA.Value = vbUnchecked And _
       chkRatios.Value = vbUnchecked And _
       chkTabSec.Value = vbUnchecked And _
       chkTipAsto.Value = vbUnchecked And _
       chkTipDoc.Value = vbUnchecked And _
       chkTipMon.Value = vbUnchecked And _
       chkCenCos.Value = vbUnchecked And _
       chkTipoCbio.Value = vbChecked Then
'       chkFlujo.Value = vbUnchecked And
        If Mensajes("Los tipos de cambio de la empresa de origen" & Salto(1) & "reemplazara a" & Salto(1) & "todos los tipos de cambio de la empresa de destino," & Salto(1) & "en el rango de fechas seleccionado" & Salto(2) & "Desea continuar...", vbQuestion + vbYesNo) = vbYes Then
            Validar = True
        Else
            Validar = False
            Exit Function
        End If
    Else
    
        If ConsultaEmpresa("EXISTEMOV", tdbcEmpresaDest.BoundText) = True Then
            Mensajes "No se puede Replicar Datos. " + Chr(13) + _
            "La empresa seleccionada ya tiene asientos registrados.", vbInformation
            Validar = False
            Exit Function
        End If
        
    End If
    
    If chkTipAsto.Value = vbChecked And chkPlan.Value = vbUnchecked Then
        Mensajes "Seleccione tambien plan de cuentas", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If

    If chkParamIni.Value = vbChecked And chkLibros.Value = vbUnchecked Then
        Mensajes "Seleccione tambien libros", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If

    If chkConfigOp.Value = vbChecked And chkPlan.Value = vbUnchecked And chkLibros.Value = vbUnchecked Then
        Mensajes "Seleccione tambien libros y plan de cuentas", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If

    If chkCtaCte.Value = vbChecked And chkPlan.Value = vbUnchecked Then
        Mensajes "Seleccione tambien plan de cuentas", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If

    If Format(dtpFechaIni.Value, "yyyyMMdd") > Format(dtpFechaFin.Value, "yyyyMMdd") Then
        Mensajes "La fecha inicial debe ser menor que la fecha final", vbOKOnly + vbExclamation
        Validar = False
        Exit Function
    End If
    
    If MsgBox("Esta seguro de importar y reemplazar los datos seleccionados, de la empresa de ORIGEN a la NUEVA empresa" & Salto(2) & "ORIGEN : " & tdbcEmpresa.Text & " - " & tdbcAnio.Text & Salto(1) & "NUEVA   : " & tdbcEmpresaDest.Text & " - " & tdbcAnioDest.Text, vbYesNo + vbQuestion, "CUIDADO") = vbNo Then
        Validar = False
        Exit Function
    End If

    Validar = True
End Function
Private Sub cmdExportar_Click()
    If Validar = True Then
        Screen.MousePointer = vbHourglass
        CargaArregloReplica
        
        Dim clsMante As New clsMantoTablas
        clsMante.InicializaClase
        clsMante.BeginTrans
        Dim i As Integer
        Dim lArrPlantCO() As Variant                 ' Arreglo para los mantenimientos
        ReDim lArrPlantCO(3) As Variant
        lArrPlantCO(0) = tdbcEmpresa.BoundText       ' Empresa
        lArrPlantCO(1) = tdbcAnio.Text               ' Anio
        lArrPlantCO(2) = gsBD
'        For i = 0 To 28
'            Debug.Print "'" & lArrMnt(i) & "',"
'        Next
         
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_UpdPeriodoPlantillaEEFF", lArrPlantCO(), True) = False Then
            Debug.Print "No se actualizo..."
        End If
        
        Set clsMante = New clsMantoTablas
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReplicarEmpresa", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido.", vbOKOnly + vbInformation
            clsMante.CancelTrans
            clsMante.FinalizaClase
            Screen.MousePointer = vbNormal
            Exit Sub
        Else
            clsMante.CommitTrans
            clsMante.FinalizaClase
            Mensajes "El proceso ha concluido.", vbOKOnly + vbInformation
        End If
        
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub cmdRefresh_Click()
    LlenaEmpresas
    DoEvents
    VerificaComboDestino
End Sub

Private Sub cmdRefreshDestino_Click()
    LlenaEmpresasDest
    DoEvents
    VerificaComboDestino
End Sub

Private Sub cmdReplicar_Click()
'    Dim respuesta As String
'    Dim añoTrabajo As String
'    If tdbgCostos.Columns(0).Value = "" Then
'        Mensajes "No existe ninguna empresa para la replicación", vbYesNo + vbCritical
'        Exit Sub
'    End If
'    ' *** Validar q empresa seleccionada no tenga movimientos
'    If ConsultaEmpresa("EXISTEMOV", tdbgCostos.Columns(0).Value) = True Then
'        Mensajes "No se puede Replicar Datos. " + Chr(13) + _
'        "La empresa seleccionada ya tiene asientos registrados.", vbInformation
'        Exit Sub
'    End If
'
'        ' *** Preguntar si se quiere replicar de verdad
'    respuesta = MsgBox("Desea replicar datos a esta empresa", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Replicar Datos")
'    If respuesta = vbYes Then
'        ' *** Buscar el valor minimo de la empresa o el unico valor
'        añoTrabajo = ConsultaEmpresaAño("ANIOMINIMO", tdbgCostos.Columns(0).Value)
'        Call Replicar(tdbgCostos.Columns(0).Value, añoTrabajo)
'
'        '*** Replicar Plan de Cuentas***
'        Call CargaPlanCTA
'''''        For i = 0 To rs_plan_cta.RecordCount
'''''
'''''
'''''
'''''
'''''        Next
'''''        '................................
'
'
'
'
'    End If
'    ' ***
End Sub

Private Sub Replicar(empresa As String, año As String)
'    Dim clsMante As clsMantoTablas
'
'    Set clsMante = New clsMantoTablas
'    ' *** Grabando Empresa
'    On Local Error GoTo ErrorEjecucion
'    ReDim lArrMnt(2) As Variant
'    lArrMnt(0) = empresa        ' EMPRESA
'    lArrMnt(1) = año            ' AÑO
'    lArrMnt(2) = gsUsuario      ' USUARIO
'    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReplicarEmpresa", lArrMnt(), True) = False Then
'        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'        Exit Sub
'    End If
'    ' ***
'
'    Mensajes "Los datos se replicaron con exito...", vbInformation
'    cmdReplicar.Enabled = False
'    Exit Sub
'ErrorEjecucion:
'    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then

        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 200
            .Height = Me.Height - .Top - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 400
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

Private Sub SSTCentroCosto_Click(PreviousTab As Integer)

    If SSTCentroCosto.Tab <> 0 Then
        tbrOpciones.Buttons(1).Enabled = False
        tbrOpciones.Buttons(5).Enabled = False
        tbrOpciones.Buttons(4).Enabled = False
        
        Me.tdbcEmpresa.Enabled = True
        Me.tdbcAnio.Enabled = True
    Else
        tbrOpciones.Buttons(5).Enabled = True
        If gsAdmin = "0" Then
           tbrOpciones.Buttons(1).Enabled = False
           tbrOpciones.Buttons(4).Enabled = False
           cmdExportar.Enabled = False
        Else
           tbrOpciones.Buttons(1).Enabled = True
           tbrOpciones.Buttons(4).Enabled = True
           cmdExportar.Enabled = True
        End If
        
    End If
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    
    If SSTCentroCosto.Tab = 2 Then
        'tdbcEmpresa.Enabled = False
        'tdbcAnio.Enabled = False

        If chkTipoCbio.Value = vbChecked Then
            dtpFechaIni.Enabled = True
            dtpFechaFin.Enabled = True
        Else
            dtpFechaIni.Enabled = False
            dtpFechaFin.Enabled = False
        End If
        
        VerificaComboDestino
        
    End If
    
End Sub

Private Sub VerificaComboDestino()

    If tdbcEmpresaDest.BoundColumn = "" Then
        tdbcEmpresaDest.Enabled = True
        tdbcAnioDest.Enabled = True
        tdbcEmpresaDest.BoundText = "< Ninguno >"
        tdbcAnioDest.BoundText = "< Ninguno >"
        
    Else
        tdbcEmpresaDest.Enabled = True
        tdbcAnioDest.Enabled = True
    End If
    
    tdbcEmpresaDest.ReBind
    tdbcAnioDest.ReBind
    DoEvents
    
End Sub
Private Sub SSTCentroCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SSTCentroCosto.Tab = 1 Then
        pSetFocus tdbtAnio
    End If
End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo
                SSTCentroCosto.TabEnabled(2) = False
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                SSTCentroCosto.TabEnabled(2) = True
        Case 4: Eliminar
        Case 5: Editar
                SSTCentroCosto.TabEnabled(2) = False
        Case 7
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    SSTCentroCosto.TabEnabled(2) = True
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
            End If
    End Select
End Sub

Private Sub ProcesoEliminarEmpresa()
    Dim clsMante As New clsMantoTablas
    ReDim lArrMnt(1) As Variant
    Dim lArrMnt2(2) As Variant
    lArrMnt(0) = tdbgCostos.Columns(0).Value
    
'    clsMante.InicializaClase
'    clsMante.BeginTrans
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_EliminarEmpresa", lArrMnt(), False, True, True) = False Then
        Mensajes "Intenete nuevamente para continuar el proceso.", vbOKOnly + vbInformation
        
'        clsMante.CommitTrans
        clsMante.FinalizaClase
        
        Exit Sub
    Else
        lArrMnt2(0) = "ELIMINAR"
        lArrMnt2(1) = tdbgCostos.Columns(0)
        lArrMnt2(2) = "001"
    
        gsImportacion = True
        DoEvents
            
            
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spSGM_EMPSOFT", lArrMnt2(), False, True, True) = False Then
            Mensajes "El proceso no ha concluido.", vbOKOnly + vbInformation
            
'            clsMante.CancelTrans
            clsMante.FinalizaClase
            gsImportacion = False
            Exit Sub
        Else
'            clsMante.CommitTrans
            clsMante.FinalizaClase
    
            Call CargaTabla
            Mensajes "La empresa fue eliminada.", vbOKOnly + vbInformation
            
        End If
    End If
    
    gsImportacion = False
  
    Set clsMante = Nothing
    
    Call Cancelar
    Call CargaTabla

End Sub


Private Sub Eliminar()

    If tdbgCostos.Columns(0).Value = gsEmpresa Then
        Mensajes "Usted quiere eliminar la empresa actual del sistema" & Salto(1) & "Con nombre : " & gsEmpresaNom & Salto(2) & "Para eliminarla cambie a una empresa diferente e intente nuevamente esta opcion", vbOKOnly + vbInformation
        Exit Sub
    End If


    Dim mensaje As String
    mensaje = MsgBox("Desea eliminar la empresa seleccionada " & Salto(2) & "EMPRESA :" & tdbgCostos.Columns(1), vbYesNo + vbExclamation, gsNombreModulo)
    
    If mensaje = vbYes Then
        If ConsultaEmpresa("EXISTEMOV", tdbgCostos.Columns(0).Value) = True Then
            If MsgBox("Desea eliminar la empresa " & tdbgCostos.Columns(1) & Salto(1) & "Toda la información que tiene se perderá", vbYesNo + vbQuestion, "Primera Confirmación") = vbYes Then
                If MsgBox("Esta seguro de eliminar la empresa " & tdbgCostos.Columns(1), vbYesNo + vbExclamation, "Segunda Confirmación") = vbYes Then
                    'ELIMINAR
                    Screen.MousePointer = vbHourglass
                    Call ProcesoEliminarEmpresa
                    Screen.MousePointer = vbNormal
                End If
            End If
        Else
            'ELIMINAR
            Screen.MousePointer = vbHourglass
            Call ProcesoEliminarEmpresa
            Screen.MousePointer = vbNormal
        End If
    End If
        
End Sub

Private Sub Editar()

    Call CargaDatosRegistro
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        
        Me.tdbtNombreLargo.Enabled = True
        Me.tdbtNombreCorto.Enabled = True
        Me.tdbtDireccion.Enabled = True
        Me.tdbtRuc.Enabled = True
        Me.tdbtTelefono.Enabled = True
        
        Call TabMantenimiento(True)
        pSetFocus tdbtNombreLargo
    Else
        lRegElim = False
    End If
    lblanio.Visible = False
    tdbtAnio.Visible = False
    tdbcSucursal.Enabled = True
    cmdEmpresas.Visible = False
    
    Me.tdbcEmpresa.Enabled = True
    Me.tdbcAnio.Enabled = True
    Me.tdbcEmpresaDest.Enabled = True
    Me.tdbcAnioDest.Enabled = True
    
    dtpFechaIni.Enabled = False
    dtpFechaFin.Enabled = False
    
End Sub

Private Sub ManNuevo()

    If LicEmpresa = False Then
        Mensajes "No se pueden registrar mas empresas. Consulte a su Proveedor", vbInformation
        Exit Sub
    End If
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)

    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    
    tbrOpciones.Buttons(1).Enabled = False
    
    tdbtCodigo = numeroEmpresa
    
    tdbtNombreCorto.Enabled = True
    tdbtNombreLargo.Enabled = True
    tdbtDireccion.Enabled = True
    tdbtRuc.Enabled = True
    tdbtTelefono.Enabled = True
    
    pSetFocus tdbtCodigo
    cmdEmpresas.Visible = True
    cmdEmpresas.Enabled = True
    lblanio.Visible = True
    tdbtAnio.Visible = True
    pSendKeys "{Enter}"
    tdbcSucursal.BoundText = Me.tdbgCostos.Columns(2).Value
    tdbcSucursal.Enabled = True
    DoEvents
    pSetFocus tdbtNombreLargo
    
    tdbtCodigo.ReadOnly = False
    tdbtCodigo.Enabled = True
    
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    SSTCentroCosto.TabEnabled(1) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    If Valor = True Then SSTCentroCosto.Tab = 1
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
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
    
    
    If gsAdmin = "0" Then
       tbrOpciones.Buttons(1).Enabled = False
       tbrOpciones.Buttons(4).Enabled = False
    Else
       tbrOpciones.Buttons(1).Enabled = True
       tbrOpciones.Buttons(4).Enabled = True
    End If
    
    pSetFocus tdbgCostos
End Sub

Private Sub CargaArregloReplica()

    ReDim lArrMnt(30) As Variant
    lArrMnt(0) = tdbcEmpresaDest.BoundText
    lArrMnt(1) = tdbcAnioDest.Text
    lArrMnt(2) = tdbcEmpresa.BoundText
    lArrMnt(3) = tdbcAnio.Text
    lArrMnt(4) = gsUsuario
    
    lArrMnt(5) = CE(chkPlan.Value)
    lArrMnt(6) = CE(chkCenCos.Value)
    lArrMnt(7) = CE(chkLibros.Value)
    
    lArrMnt(8) = CE(chkEntidades.Value)
    lArrMnt(9) = CE(chkEntidades.Value)
    lArrMnt(10) = CE(chkEntidades.Value)
    
    lArrMnt(11) = CE(chkTipDoc.Value)
    lArrMnt(12) = CE(chkTabSec.Value)
    lArrMnt(13) = CE(chkTipMon.Value)
    lArrMnt(14) = CE(chkTipoCbio.Value)
    lArrMnt(15) = CE(chkBancos.Value)
    lArrMnt(16) = CE(chkCtaCte.Value)
    lArrMnt(17) = CE(chkTipAsto.Value)
    lArrMnt(18) = CE(chkPlanCONA.Value)
    
    lArrMnt(19) = CE(chkParamIni.Value)
    lArrMnt(20) = CE(chkParamIni.Value)
    
    lArrMnt(21) = CE(chkConfigOp.Value)
    lArrMnt(22) = CE(chkConfigOp.Value)
    
    lArrMnt(23) = CE(chkRatios.Value)
    lArrMnt(24) = CE(dtpFechaIni.Value)
    lArrMnt(25) = CE(dtpFechaFin.Value)
    
    lArrMnt(26) = NE(tdbnDigitos.Value)
    
    lArrMnt(27) = CE(0) 'CE(chkFlujo.Value)
    lArrMnt(28) = CE(chkDetalleEnt.Value)
    
    'Se agrego valores a las variables BDOrigen y BDDestino (Nombre Base de Datos)
    lArrMnt(29) = Trim$(gsBD)
    lArrMnt(30) = Trim$(gsBD)
    
End Sub

Private Sub Grabar()
    Dim clsMante As clsMantoTablas
    Dim condicion As Boolean
    
    If validarDatos = False Then Exit Sub
    Set clsMante = New clsMantoTablas
    
    If Not fValidaRUC() Then
       pSetFocus tdbtRuc
       Exit Sub
    End If
    
    ' ****************** Grabando Empresa *****************
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt
    condicion = True
    
    If lblMante = "NUEVO REGISTRO" Then
        Call EscribirLog("Se esta creando la empresa " & tdbtCodigo.Text & " " & tdbtNombreLargo.Text, Me.Name)
    Else
        Call EscribirLog("Se ha modificado los datos de la empresa " & tdbtCodigo.Text & " " & tdbtNombreLargo.Text, Me.Name)
    End If
    
    clsMante.InicializaClase
    clsMante.BeginTrans
    
    If lTipoMnt = "INSERTAR" Then condicion = False
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEmpresa", lArrMnt(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        
        clsMante.CancelTrans
        clsMante.FinalizaClase
        Exit Sub
    End If
    ' *****************************************************
    clsMante.CommitTrans
    ' *****************************************************
    
        CargaArregloMntEMPSOFT
        condicion = True
        If lTipoMnt = "INSERTAR" Then condicion = False
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spSGM_EMPSOFT", lArrMnt(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            
            clsMante.CancelTrans
            clsMante.FinalizaClase
            
            Exit Sub
        End If
        ' *****************************************************

    clsMante.CommitTrans
    clsMante.FinalizaClase
   
    Set clsMante = Nothing
    
    Call Cancelar
    CargaDatosIniciales gsEmpresa
    
    CargaTabla
    
    ' *** Buscar la empresa creada y posicionarse alli ****
    Dim Valor As Integer
    Valor = BuscarCadRs(tdbtCodigo, lrsTabla, 1)
    If Valor = 0 Then lrsTabla.MoveFirst
    ' *****************************************************
    
    Mensajes "Los datos se grabaron con exito...", vbInformation
    
    Call EscribirLog("Finalizo la creacion de la empresa " & tdbtCodigo.Text & " " & tdbtNombreLargo.Text, Me.Name)
    
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
    
    If lTipoMnt = "INSERTAR" Then
        Mensajes "Defina la configuración inicial de la empresa", vbOKOnly + vbExclamation
        SSTCentroCosto.Tab = 2
        DoEvents
        cmdRefresh_Click
        DoEvents
        
        tdbcEmpresaDest.BoundText = tdbtCodigo.Text
        tdbcAnioDest.BoundText = tdbtAnio.Text
        tdbcAnioDest.Enabled = True
        tdbcEmpresaDest.Enabled = True
        
    End If
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    Call EscribirLog("Error al crear la empresa " & tdbtCodigo.Text & " " & tdbtNombreLargo.Text & ", [" & Err.Description & "]", Me.Name)
    
End Sub

Private Sub CargaDatosIniciales(empresa As String)
    Dim sqlSp As String
    Dim arrDatos() As Variant
    Dim rsArreglo  As ADODB.Recordset
    Set rsArreglo = New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    
    sqlSp = "select EMP_CNOMBRELARGO, emp_cnumruc, ISNULL(Emp_Bymoneda, 0) AS Emp_Bymoneda from EMPRESA where emp_ccodigo='" & empresa & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
        If rsArreglo.State = 1 Then
           gsRUC = CE(rsArreglo("emp_cnumruc"))
           gsEmpresaNom = CE(rsArreglo("EMP_CNOMBRELARGO"))
           'Cargar el valor del Campo Emp_Bymoneda
           gintBiMoneda = IIf(IsNull(CE(rsArreglo("Emp_Bymoneda"))), 0, CE(rsArreglo("Emp_Bymoneda")))
        End If
    End If
    
    Set clDatos = Nothing
    CerrarRecordSet rsArreglo
End Sub
Private Function CargaPlanCTA() As ADODB.Recordset

    
End Function
    
Private Function validarDatos() As Boolean
    validarDatos = False
    
    If lTipoMnt = "INSERTAR" Then
        If MsgBox("Desea crear la empresa: " & tdbtNombreLargo & Salto(2) & "Con el periodo contable : " & tdbtAnio.Value, vbOKCancel + vbQuestion) = vbCancel Then
           Exit Function
        End If
    End If
    
    ' *** Validar q los datos necesarios esten ingresados
    If TextoLleno(Me.tdbtCodigo, "Codigo") = False Then Exit Function
    If TextoLleno(Me.tdbtNombreLargo, "Nombre Largo") = False Then Exit Function
   
    If Not fValidaRUC() Then
       pSetFocus tdbtRuc
       Exit Function
    End If
    
    ' ***
    validarDatos = True
End Function

Private Sub CargaArregloMntEMPSOFT()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(2) As Variant
    lArrMnt(0) = CE(lTipoMnt)           ' Accion
    lArrMnt(1) = CE(tdbtCodigo)         ' Empresa
    lArrMnt(2) = "001"                  ' codigo de SW
End Sub

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(10) As Variant
    lArrMnt(0) = CE(lTipoMnt)           ' Accion
    lArrMnt(1) = CE(tdbtCodigo)         ' Empresa
    lArrMnt(2) = CE(tdbtNombreLargo)    ' nombre
    lArrMnt(3) = CE(tdbtNombreCorto)    ' Nombre corto
    lArrMnt(4) = CE(tdbtDireccion)      ' Direccion
    lArrMnt(5) = CE(tdbtRuc)            ' Ruc
    lArrMnt(6) = CE(tdbtTelefono)       ' Telefono
    lArrMnt(7) = CE(Me.tdbtAnio)        ' Año (para replicarlo junto a sus periodos)
    lArrMnt(8) = CE(Me.tdbcSucursal.BoundText)  ' Sucursal
    lArrMnt(9) = gsUsuario          ' Usuario
    'Asignando el valor del control al parametro Emp_Bymoneda
    lArrMnt(10) = CE(Me.chkBiMoneda.Value)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case vbKeyF1
        
    End Select
    ' ***
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Dim sqlcombos As String
    CodigoEmp = ""
    
    Centrar_form Me

    Call CargaTabla

    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.Tab = 0
    SSTCentroCosto.TabEnabled(1) = False
    tdbtAnio = Year(FechaServidor)

    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA "
    sqlcombos = sqlcombos + "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Tab_cTabla = '031' ORDER BY Tab_cCodigo"
    LlenarComboAddItem tdbcSucursal, sqlcombos
    
    ActivarControl dtpFechaIni, False
    ActivarControl dtpFechaFin, False
    
    DoEvents
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    
    DoEvents
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdExportar.Enabled = False
    Else
        Me.cmdExportar.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    LlenaEmpresas
    DoEvents
    LlenaEmpresasDest
    
    If gsAdmin = "0" Then
       tbrOpciones.Buttons(1).Enabled = False
       tbrOpciones.Buttons(4).Enabled = False
       cmdExportar.Enabled = False
    Else
       tbrOpciones.Buttons(1).Enabled = True
       tbrOpciones.Buttons(4).Enabled = True
       cmdExportar.Enabled = True
    End If
    
End Sub

Private Sub tdbcEmpresaDest_ItemChange()
    Dim sqlCadena As String
    Dim posicion As Integer
    
    On Local Error GoTo ErrorEjecucion
    sqlCadena = "SELECT PAN_CANIO, PAN_CANIO as 'PER_CPERIODO' FROM CNT_ANIO WHERE EMP_CCODIGO = '" & tdbcEmpresaDest.BoundText & "' "
    sqlCadena = sqlCadena + "ORDER BY PAN_CANIO "
    
    ComboArreglo lArrAnio, tdbcAnioDest, sqlCadena
    
    tdbcAnioDest.Bookmark = 0
    Exit Sub
ErrorEjecucion:
    Exit Sub
End Sub

Private Sub tdbcEmpresa_ItemChange()
    Dim sqlCadena As String
    Dim posicion As Integer
    
    On Local Error GoTo ErrorEjecucion
    sqlCadena = "SELECT PAN_CANIO, PAN_CANIO as 'PER_CPERIODO' FROM CNT_ANIO WHERE EMP_CCODIGO = '" & tdbcEmpresa.BoundText & "' "
    sqlCadena = sqlCadena + "ORDER BY PAN_CANIO "
    
    ComboArreglo lArrAnio, tdbcAnio, sqlCadena
    
    tdbcAnio.BoundText = gsAnio
    'tdbcAnio.Enabled = False
    
    LlenaEmpresasDest
    Exit Sub
ErrorEjecucion:
    Exit Sub
End Sub

Private Sub LlenaEmpresasDest()
    Dim sqlCadena As String
    
    
    sqlCadena = "spCN_GrabaEmpresa 'SEL_DESTINO','" & gsEmpresa & "','','','','','','" & gsAnio & "','001','" & gsUsuario & "',''"
    
    tdbcEmpresaDest.Columns(2).Visible = False
    tdbcEmpresaDest.Columns(3).Visible = False
   
    LlenarComboAddItem tdbcEmpresaDest, sqlCadena
    tdbcEmpresaDest.BoundText = gsEmpresa
    
    tdbcEmpresaDest.Enabled = True
End Sub

Private Sub LlenaEmpresas()
    Dim sqlCadena As String
    
    sqlCadena = "spCN_GrabaEmpresa 'SEL_ALL','" & gsEmpresa & "','','','','','','" & gsAnio & "','001','" & gsUsuario & "',''"
    
    tdbcEmpresa.Columns(2).Visible = False
    tdbcEmpresa.Columns(3).Visible = False
   
    LlenarComboAddItem tdbcEmpresa, sqlCadena
    tdbcEmpresa.BoundText = gsEmpresa
    
    'Me.tdbcEmpresa.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaEmpresa 'SEL_ALL', '', '', '', '', '', '', '', '','" & gsUsuario & "','' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then
        ' *** Llenar grilla con el RecordSet
        lrsTabla.Sort = "Emp_cCodigo"
        tdbgCostos.DataSource = lrsTabla
        ' ***
        Exit Sub
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    sqlSp = "spCn_GrabaEmpresa 'SEL_REG', '" & tdbgCostos.Columns(0).Value & "', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro el registro. Probablemente eliminado desde otra sesion", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    ' *** Asignando Datos de la Cuenta
    tdbtCodigo = CE(rsArreglo!Emp_cCodigo)
    tdbtNombreLargo = CE(rsArreglo!EMP_CNOMBRELARGO)
    tdbtNombreCorto = CE(rsArreglo!Emp_cNombreCorto)
    tdbtDireccion = CE(rsArreglo!Emp_cDireccion)
    tdbtRuc = CE(rsArreglo!Emp_cNumRuc)
    tdbtTelefono = CE(rsArreglo!Emp_cTelefono)
    Me.tdbcSucursal.BoundText = CE(rsArreglo!Emp_cCodSuc)
    
    'Carga el Valor del Campo Emp_Bymoneda
    Me.chkBiMoneda.Value = CE(rsArreglo!emp_bymoneda)
    'Si el valor del Campo Emp_Bymoneda es 1 se bloquea el control para que no pueda ser modificado
'    If (CInt(CE(rsArreglo!emp_bymoneda)) = 1) Then
'        Me.chkBiMoneda.Enabled = False
'    Else
'        Me.chkBiMoneda.Enabled = True
'    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(0) As String
    Dim i As Integer
    
    If lrsTabla Is Nothing Then Exit Sub
    cadena = ""
    If Trim(tdbtDescripcionBus) <> "" Then filtros(0) = "Emp_cNombreLargo like '*" & tdbtDescripcionBus & "*'"
    For i = 0 To 0
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

Private Sub tdbtAnio_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbtAnio.ReadOnly = True
    Else
        tdbtAnio.ReadOnly = False
    End If
    ' ***
End Sub

Private Sub tdbtCodigo_Change()
    If CE(tdbtCodigo.Text) = "" Then
       CodigoEmp = ""
       
       BloquearControles True
       LimpiarControles
       
    End If
    
    If CE(tdbtCodigo.Text) <> CodigoEmp Then
       BloquearControles True
       LimpiarControles
    End If
    
End Sub

Private Sub tdbtCodigo_LostFocus()
        
        If lTipoMnt = "INSERTAR" Then
            If CE(tdbtCodigo) = "000" Then
                Mensajes "Ingrese otro codigo. Es un codigo de Sistema", vbInformation
                pSetFocus tdbtCodigo
                tdbtCodigo = ""
                Exit Sub
            End If

            If ExisteCodigo(tdbtCodigo) = True Then
                Mensajes "El codigo de empresa ya existe. Verifique...", vbInformation
                pSetFocus tdbtCodigo
                tdbtCodigo = ""
            End If
        End If

End Sub

Private Function ExisteCodigo(Valor As String) As Boolean
    ' *** Verificar q codigo exista
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    ExisteCodigo = False
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaEmpresa 'SEL_REG', '" & Valor & "', '','', '', '', '', '','','','' "
        arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteCodigo = True
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Sub tdbtDescripcionBus_Change()
    
    If gsKey = 219 Then
       tdbtDescripcionBus = Replace(tdbtDescripcionBus, "'", "")
       tdbtDescripcionBus.SelStart = Len(tdbtDescripcionBus)
    End If
    
    Call FiltrarRecordSet
End Sub

Private Sub tdbtCodigo_GotFocus()
    ' *** Si es modificar no editar
    '
    If lTipoMnt = "EDITAR" Then
        tdbtCodigo.ReadOnly = True
    Else
        tdbtCodigo.ReadOnly = False
    End If
    
    ' ***
End Sub

Private Function ConsultaEmpresa(Tipo As String, empresa As String) As Boolean
    Dim rsDatos As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    sqlDatos = "spCn_GrabaEmpresa '" & Tipo & "', '" & empresa & "', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsDatos(0).Value = 0 Then
        ConsultaEmpresa = False
    Else
        ConsultaEmpresa = True
    End If
    Call CerrarRecordSet(rsDatos)
End Function

Private Function ConsultaEmpresaAño(Tipo As String, empresa As String) As String
    Dim rsDatos As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    ConsultaEmpresaAño = ""
    sqlDatos = "spCn_GrabaEmpresa '" & Tipo & "', '" & empresa & "', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsDatos.State = 0 Then Exit Function
    ConsultaEmpresaAño = rsDatos(0).Value
    Call CerrarRecordSet(rsDatos)
End Function

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

Private Sub tdbtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDireccion = Replace(tdbtDireccion, "'", "")
       tdbtDireccion.SelStart = Len(tdbtDireccion)
    End If
End Sub

Private Sub tdbtNombreCorto_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtNombreCorto = Replace(tdbtNombreCorto, "'", "")
       tdbtNombreCorto.SelStart = Len(tdbtNombreCorto)
    End If
End Sub

Private Sub tdbtNombreLargo_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtNombreLargo = Replace(tdbtNombreLargo, "'", "")
       tdbtNombreLargo.SelStart = Len(tdbtNombreLargo)
    End If
End Sub

Private Sub tdbtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not fValidaRUC() Then
          pSetFocus tdbtRuc
        Else
          pSetFocus tdbtTelefono
       End If
    End If
End Sub

Private Function fValidaRUC() As Boolean
    fValidaRUC = False
    
    If Trim(tdbtRuc) <> "" Then
        If Len(Trim(tdbtRuc)) <> 11 And Me.tdbtRuc.Enabled = True Then
            Mensajes "Numero de digitos de Ruc debe ser igual a 11. Verificar.. ", vbInformation
            Exit Function
        End If
        
        ' Verifica si es correcto el nro de RUC
        If Not fValidarNroRuc(tdbtRuc) Then
           MsgBox "El RUC no es válido", vbInformation
           Exit Function
        End If
    End If
    
    fValidaRUC = True
End Function

Private Sub tdbtRuc_LostFocus()
If Not fValidarNroRuc(tdbtRuc) And CE(tdbtRuc.Text) <> "" And SSTCentroCosto.TabEnabled(2) = False Then
   MsgBox "El RUC no es valido", vbInformation
   pSetFocus tdbtRuc
   Exit Sub
End If
End Sub

Private Sub cmdEmpresas_Click()
       Call LlamaBuscar(frmBuscador, "tdbtCodigo", "tdbtCodigo", "Empresas", Me, gsPeriodo)
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal param4 As String, ByVal param5 As String, ByVal param6 As String)
    CodigoEmp = CE(param0)
    tdbtCodigo.Text = CE(param0)
    tdbtNombreLargo.Text = CE(param1)
    tdbtNombreCorto.Text = CE(param2)
    tdbtDireccion.Text = CE(param3)
    tdbtRuc.Text = CE(param4)
    tdbtTelefono.Text = CE(param5)
    tdbcSucursal.BoundText = CE(param6)
    Unload frmBuscador
    
    BloquearControles False
End Sub

Private Sub BloquearControles(Valor As Boolean)
    tdbtNombreLargo.Enabled = Valor
    tdbtNombreCorto.Enabled = Valor
    tdbtDireccion.Enabled = Valor
    tdbtRuc.Enabled = Valor
    tdbtTelefono.Enabled = Valor
    tdbcSucursal.Enabled = Valor
End Sub

Private Sub LimpiarControles()
    tdbtNombreLargo.Text = ""
    tdbtNombreCorto.Text = ""
    tdbtDireccion.Text = ""
    tdbtRuc.Text = ""
    tdbtTelefono.Text = ""
End Sub
