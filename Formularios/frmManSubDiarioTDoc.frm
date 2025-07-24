VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmManSubDiarioTDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros Iniciales"
   ClientHeight    =   7890
   ClientLeft      =   3405
   ClientTop       =   3600
   ClientWidth     =   11655
   Icon            =   "frmManSubDiarioTDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11655
   Begin TabDlg.SSTab SSTabParam 
      Height          =   7395
      Left            =   60
      TabIndex        =   0
      Top             =   405
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Param. Iniciales de Contabilidad"
      TabPicture(0)   =   "frmManSubDiarioTDoc.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         ForeColor       =   &H8000000F&
         Height          =   7125
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   11205
         Begin TDBText6Ctl.TDBText txtDescripcionNIF 
            Height          =   315
            Left            =   1980
            TabIndex        =   65
            Top             =   2820
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":0EE6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":0F52
            Key             =   "frmManSubDiarioTDoc.frx":0F70
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
         Begin TDBText6Ctl.TDBText txtIdNIF 
            Height          =   315
            Left            =   1635
            TabIndex        =   64
            Top             =   2820
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":0FB4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1020
            Key             =   "frmManSubDiarioTDoc.frx":103E
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
         Begin TDBText6Ctl.TDBText txtDescripTransAutomatico 
            Height          =   315
            Left            =   7560
            TabIndex        =   61
            Top             =   2420
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1082
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":10EE
            Key             =   "frmManSubDiarioTDoc.frx":110C
            BackColor       =   16249284
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
         Begin TDBText6Ctl.TDBText txtIdTransAutomatica 
            Height          =   315
            Left            =   7215
            TabIndex        =   60
            Top             =   2420
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1150
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":11BC
            Key             =   "frmManSubDiarioTDoc.frx":11DA
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
         Begin TDBText6Ctl.TDBText txtDescripcionTransferencia 
            Height          =   315
            Left            =   1980
            TabIndex        =   59
            Top             =   2420
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":121E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":128A
            Key             =   "frmManSubDiarioTDoc.frx":12A8
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
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
         Begin TDBText6Ctl.TDBText txtIdTransferencia 
            Height          =   315
            Left            =   1635
            TabIndex        =   58
            Top             =   2420
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":12EC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1358
            Key             =   "frmManSubDiarioTDoc.frx":1376
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
         Begin VB.Frame Frame4 
            Caption         =   "LE Simplificado"
            Height          =   615
            Left            =   7800
            TabIndex        =   54
            Top             =   6360
            Width           =   3075
            Begin VB.CheckBox chkLECompra 
               Caption         =   "LE Compra"
               Height          =   255
               Left            =   1680
               TabIndex        =   56
               Tag             =   "_"
               Top             =   240
               Width           =   1092
            End
            Begin VB.CheckBox chkLEVenta 
               Caption         =   "LE Venta"
               Height          =   255
               Left            =   240
               TabIndex        =   55
               Tag             =   "_"
               Top             =   240
               Width           =   1092
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Libros Electrónicos"
            Height          =   600
            Left            =   7800
            TabIndex        =   53
            Top             =   4680
            Width           =   3075
            Begin VB.CheckBox chkRVIE 
               Caption         =   "SIRE"
               Height          =   255
               Left            =   1860
               TabIndex        =   67
               Tag             =   "_"
               Top             =   240
               Width           =   1155
            End
            Begin VB.CheckBox chkVersionLE 
               Caption         =   "PLE"
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Tag             =   "_"
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame fraDiarioSimplificado 
            Caption         =   "Diario Formato Simplificado"
            Height          =   945
            Left            =   7800
            TabIndex        =   47
            Top             =   5400
            Width           =   3075
            Begin VB.CheckBox chkDiarioSimplificado 
               Alignment       =   1  'Right Justify
               Caption         =   "Habilitar"
               Height          =   195
               Left            =   330
               TabIndex        =   48
               Tag             =   "_"
               Top             =   285
               Width           =   2250
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnDigReporte 
               Height          =   285
               Left            =   1935
               TabIndex        =   49
               Tag             =   "_"
               Top             =   525
               Width           =   660
               _Version        =   65536
               _ExtentX        =   1164
               _ExtentY        =   503
               Calculator      =   "frmManSubDiarioTDoc.frx":13BA
               Caption         =   "frmManSubDiarioTDoc.frx":13DA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManSubDiarioTDoc.frx":143E
               Keys            =   "frmManSubDiarioTDoc.frx":145C
               Spin            =   "frmManSubDiarioTDoc.frx":14B4
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,##0"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   7798785
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "N° Digitos Reporte"
               Height          =   195
               Index           =   6
               Left            =   375
               TabIndex        =   50
               Top             =   600
               Width           =   1320
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo de Plan de Cuentas"
            Height          =   960
            Left            =   7800
            TabIndex        =   44
            Top             =   3555
            Width           =   3075
            Begin VB.OptionButton optPlan 
               Caption         =   "Empresarial - PCGE"
               Height          =   330
               Index           =   1
               Left            =   540
               TabIndex        =   46
               Top             =   540
               Width           =   1950
            End
            Begin VB.OptionButton optPlan 
               Caption         =   "Revisado - PCGR"
               Height          =   330
               Index           =   0
               Left            =   540
               TabIndex        =   45
               Top             =   270
               Value           =   -1  'True
               Width           =   1950
            End
         End
         Begin VB.CheckBox chkCondicion 
            Caption         =   "Trabajar con Caja Ingresos y Caja Egresos"
            Height          =   225
            Left            =   5820
            TabIndex        =   3
            Top             =   1215
            Width           =   4320
         End
         Begin VB.Frame Frame5 
            Height          =   60
            Left            =   60
            TabIndex        =   2
            Top             =   3360
            Width           =   11130
         End
         Begin TDBText6Ctl.TDBText tdbtDescripVentas 
            Height          =   315
            Left            =   1980
            TabIndex        =   4
            Top             =   300
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":14DC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1548
            Key             =   "frmManSubDiarioTDoc.frx":1566
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDiarioVentas 
            Height          =   315
            Left            =   1635
            TabIndex        =   5
            Top             =   300
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":15A8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1614
            Key             =   "frmManSubDiarioTDoc.frx":1632
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
         Begin TDBText6Ctl.TDBText tdbtDiarioCompras 
            Height          =   315
            Left            =   1635
            TabIndex        =   6
            Top             =   720
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1674
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":16E0
            Key             =   "frmManSubDiarioTDoc.frx":16FE
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
         Begin TDBText6Ctl.TDBText tdbtDiarioDif 
            Height          =   315
            Left            =   1635
            TabIndex        =   7
            Top             =   2025
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1740
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":17AC
            Key             =   "frmManSubDiarioTDoc.frx":17CA
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
         Begin TDBText6Ctl.TDBText tdbtDiarioCaja 
            Height          =   315
            Left            =   1635
            TabIndex        =   8
            Top             =   1155
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":180C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1878
            Key             =   "frmManSubDiarioTDoc.frx":1896
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
         Begin TDBText6Ctl.TDBText tdbtDiarioCajaIng 
            Height          =   315
            Left            =   7215
            TabIndex        =   9
            Top             =   1605
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":18D8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1944
            Key             =   "frmManSubDiarioTDoc.frx":1962
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
         Begin TDBText6Ctl.TDBText tdbtDescripCompras 
            Height          =   315
            Left            =   1980
            TabIndex        =   10
            Top             =   720
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":19A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1A10
            Key             =   "frmManSubDiarioTDoc.frx":1A2E
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDescripDif 
            Height          =   315
            Left            =   1980
            TabIndex        =   11
            Top             =   2025
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1A70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1ADC
            Key             =   "frmManSubDiarioTDoc.frx":1AFA
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDescripCaja 
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Top             =   1155
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1B3C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1BA8
            Key             =   "frmManSubDiarioTDoc.frx":1BC6
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDescripCajaIng 
            Height          =   315
            Left            =   7560
            TabIndex        =   13
            Top             =   1605
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1C08
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1C74
            Key             =   "frmManSubDiarioTDoc.frx":1C92
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDiarioCajaEgr 
            Height          =   315
            Left            =   7215
            TabIndex        =   14
            Top             =   2025
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1CD4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1D40
            Key             =   "frmManSubDiarioTDoc.frx":1D5E
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
         Begin TDBText6Ctl.TDBText tdbtDescripCajaEgr 
            Height          =   315
            Left            =   7560
            TabIndex        =   15
            Top             =   2025
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1DA0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1E0C
            Key             =   "frmManSubDiarioTDoc.frx":1E2A
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBNumber6Ctl.TDBNumber tdbnIGV 
            Height          =   285
            Left            =   6450
            TabIndex        =   20
            Top             =   5430
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":1E6C
            Caption         =   "frmManSubDiarioTDoc.frx":1E8C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1EF0
            Keys            =   "frmManSubDiarioTDoc.frx":1F0E
            Spin            =   "frmManSubDiarioTDoc.frx":1F66
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   7798785
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBText6Ctl.TDBText tdbtDiario 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   1575
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":1F8E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1FFA
            Key             =   "frmManSubDiarioTDoc.frx":2018
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
         Begin TDBText6Ctl.TDBText tdbtDescripDiario 
            Height          =   315
            Left            =   1980
            TabIndex        =   23
            Top             =   1575
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":205A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":20C6
            Key             =   "frmManSubDiarioTDoc.frx":20E4
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDiarioCierre 
            Height          =   315
            Left            =   7215
            TabIndex        =   24
            Top             =   330
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":2126
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":2192
            Key             =   "frmManSubDiarioTDoc.frx":21B0
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
         Begin TDBText6Ctl.TDBText tdbtDescripCierre 
            Height          =   315
            Left            =   7560
            TabIndex        =   25
            Top             =   330
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":21F2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":225E
            Key             =   "frmManSubDiarioTDoc.frx":227C
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDiarioApe 
            Height          =   315
            Left            =   7215
            TabIndex        =   26
            Top             =   765
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":22BE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":232A
            Key             =   "frmManSubDiarioTDoc.frx":2348
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
         Begin TDBText6Ctl.TDBText tdbtDescripApe 
            Height          =   315
            Left            =   7560
            TabIndex        =   27
            Top             =   765
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   "frmManSubDiarioTDoc.frx":238A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":23F6
            Key             =   "frmManSubDiarioTDoc.frx":2414
            BackColor       =   16249284
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   "a"
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
         Begin TDBNumber6Ctl.TDBNumber tdbnUIT 
            Height          =   285
            Left            =   6450
            TabIndex        =   21
            Top             =   6015
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":2456
            Caption         =   "frmManSubDiarioTDoc.frx":2476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":24DA
            Keys            =   "frmManSubDiarioTDoc.frx":24F8
            Spin            =   "frmManSubDiarioTDoc.frx":2550
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   7798785
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnMesCompras 
            Height          =   285
            Left            =   3825
            TabIndex        =   18
            Top             =   5355
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":2578
            Caption         =   "frmManSubDiarioTDoc.frx":2598
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":25FC
            Keys            =   "frmManSubDiarioTDoc.frx":261A
            Spin            =   "frmManSubDiarioTDoc.frx":2672
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   3080193
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnNdigitos 
            Height          =   285
            Left            =   3825
            TabIndex        =   19
            Top             =   6000
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":269A
            Caption         =   "frmManSubDiarioTDoc.frx":26BA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":271E
            Keys            =   "frmManSubDiarioTDoc.frx":273C
            Spin            =   "frmManSubDiarioTDoc.frx":2794
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,##0"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   3080193
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSDataListLib.DataCombo tdbcBaseImponible 
            Height          =   300
            Left            =   3840
            TabIndex        =   17
            Top             =   4410
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo tdbcNivel 
            Height          =   300
            Left            =   3840
            TabIndex        =   16
            Top             =   3960
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo tdbcBaseImponibleVtas 
            Height          =   300
            Left            =   3840
            TabIndex        =   52
            Top             =   4800
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label13 
            Caption         =   "Libro Ajuste NIF"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2920
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Libro TyC Automatico"
            Height          =   375
            Left            =   5520
            TabIndex        =   62
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Libro Trans-Canc"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Base Imponible x defecto Reg. Ventas"
            Height          =   270
            Index           =   7
            Left            =   105
            TabIndex        =   51
            Top             =   4815
            Width           =   2910
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Libro Cierre"
            Height          =   195
            Left            =   5820
            TabIndex        =   42
            Top             =   405
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Libro Diario"
            Height          =   195
            Left            =   225
            TabIndex        =   41
            Top             =   1620
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "IGV ( % )"
            Height          =   195
            Left            =   5550
            TabIndex        =   40
            Top             =   5445
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Caja Egresos"
            Height          =   195
            Left            =   5820
            TabIndex        =   39
            Top             =   2055
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Caja Ingreso"
            Height          =   195
            Left            =   5820
            TabIndex        =   38
            Top             =   1620
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Libro Caja"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   1215
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Libro Dif. Cambio"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   2070
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Libro Compras"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Libro Ventas"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   375
            Width           =   885
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Libro Apertura"
            Height          =   195
            Left            =   5820
            TabIndex        =   33
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label4 
            Caption         =   "Nivel asignado como Centro de costo"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   4050
            Width           =   2970
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "U.I.T."
            Height          =   195
            Index           =   2
            Left            =   5580
            TabIndex        =   31
            Top             =   6060
            Width           =   405
         End
         Begin VB.Label Label4 
            Caption         =   "Meses ant. permitidos en ingreso docs. en Libro de Compras"
            Height          =   420
            Index           =   3
            Left            =   90
            TabIndex        =   30
            Top             =   5325
            Width           =   3585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N° Digitos Cuenta Detalle, Plan cuentas"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   6015
            Width           =   2820
         End
         Begin VB.Label Label4 
            Caption         =   "Base Imponible x defecto Reg. Compras"
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   4380
            Width           =   2910
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":27BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2916
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":3132
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":328C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":3826
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":3DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":435A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":48F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":4E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":5428
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":59C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   264
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   3588
      _ExtentX        =   6324
      _ExtentY        =   476
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
Attribute VB_Name = "frmManSubDiarioTDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim lrsDiario As ADODB.Recordset
Dim lrsTdoc As ADODB.Recordset
Dim lArrMnt() As Variant
Dim lTipoMnt As String
Dim lRegElim As Boolean
Dim sqlSp As String
Dim Control As String

Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Sub pLlenarVariablesSubDiarios()
    Dim clDatos As clsMantoTablas
    Dim rs As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim bandera As Integer
    bandera = 1

    Set clDatos = New clsMantoTablas
    sqlSp = "spCNT_CONFIG_LIBROS 'BUSCARTODOS','" & gsEmpresa & "','','','','','','','','',0,'','','','','','','','','','','','','','','','" & gsAnio & "'"
    arrDatos = Array(sqlSp)
    Set rs = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rs Is Nothing Then Exit Sub
    
    tdbtDiarioVentas.Text = CE(rs!Cfl_cVentas)
    tdbtDiarioCompras.Text = CE(rs!Cfl_cCompras)
    tdbtDiario.Text = CE(rs!Cfl_cDiario)
    tdbtDiarioDif.Text = CE(rs!Cfl_cDifCam)
    tdbtDiarioCierre.Text = CE(rs!Cfl_cCierre)

    tdbtDiarioApe.Text = CE(rs!Cfl_cApertura)
    Me.txtIdTransferencia.Text = CE(rs!cfl_cTransferencia)
    Me.txtIdTransAutomatica.Text = CE(rs!Cfl_cTransAutomatico)
    
    Me.txtIdNIF.Text = CE(rs!Cfl_cAjusteNIF)
    lsLibroAjusteNIF = CE(rs!Cfl_cAjusteNIF)
    
    ''hlp20230108
    If rs!Cfl_cVersionLE = 0 And rs!Cfl_cRVIE = 0 Then
       Me.chkVersionLE.Value = 1
    ElseIf rs!Cfl_cVersionLE = 1 And rs!Cfl_cRVIE = 1 Then
       Me.chkVersionLE.Value = 1
       Me.chkRVIE.Value = 0
    Else
       Me.chkVersionLE.Value = rs!Cfl_cVersionLE
       Me.chkRVIE.Value = rs!Cfl_cRVIE 'frt_rvie
       bandera = 0
    End If
    ''hlp20230108
    If CE(CE(rs!Cfl_cCaja)) = "" Then
        chkCondicion.Value = 1
        tdbtDiarioCaja = ""
        tdbtDiarioCajaIng = CE(rs!Cfl_cCajaIngresos)
        tdbtDiarioCajaEgr = CE(rs!Cfl_cCajaEgresos)
    Else
        chkCondicion.Value = 0
        tdbtDiarioCaja = CE(rs!Cfl_cCaja)
        tdbtDiarioCajaIng = ""
        tdbtDiarioCajaEgr = ""
    End If
    lsLibroTransferenciaCancelacion = CE(rs!cfl_cTransferencia)
    lsLibroTransCancAutomatico = CE(rs!Cfl_cTransAutomatico)
    gstrVersionLE = CE(Me.chkVersionLE.Value)
    gsRVIE = CE(Me.chkRVIE.Value) 'frt_rvie
    tdbnMesCompras.Value = NE(rs!Cfl_cMesCompras)
 
    If CE(rs!Cfl_cBaseDefCom) <> "006" And CE(rs!Cfl_cBaseDefCom) <> "007" And CE(rs!Cfl_cBaseDefCom) <> "008" Then
        tdbcBaseImponible.BoundText = ""
    Else
        tdbcBaseImponible.BoundText = CE(rs!Cfl_cBaseDefCom)
    End If
    
    tdbcBaseImponibleVtas.BoundText = CE(rs!Cfl_cBaseDefVtas)
    
    tdbcNivel.BoundText = CE(rs!Cfl_cNivelCC)
    chkDiarioSimplificado.Value = vbUnchecked
    
    'chkPle.Value = IIf(IsNull(rs!Cfl_cPle) = True, 0, IIf(rs!Cfl_cPle = 0, 0, 1)) 'frt_202011

    On Error Resume Next
    gsTipoPlan = NE(rs!Cfl_cTipoPlan)
    optPlan(gsTipoPlan).Value = True
    tdbnDigReporte.Value = NE(rs!Cfl_nDigDiarioRep)
    
    chkDiarioSimplificado.Value = NE(rs!Cfl_cDiarioSimplificado)
    Me.chkLEVenta.Value = NE(rs!Cfl_cLEVenta)
    Me.chkLECompra.Value = NE(rs!Cfl_cLECompra)
    
'    Me.txtIdNIF.Text = CE(rs!CflcAjusteNIF)
'    lsLibroAjusteNIF = CE(rs!CflcAjusteNIF)
    gintLEVentaSimplificado = CInt(Me.chkLEVenta.Value)
    gintLECompraSimplificado = CInt(Me.chkLECompra.Value)
    
    Call CerrarRecordSet(rs)
    Set rs = Nothing
    ''hlp20230801
    If bandera = 1 Then
       Call Grabar(0)
    End If
    ''hlp20230801
End Sub
Private Sub LlenaOtrasConfigIniciales()
 tdbnNdigitos.Value = NE(BuscaValorEnOp("054"))
 tdbnUIT.Value = NE(BuscaValorEnOp("027"))
 tdbnIGV.Value = NE(fIgv())
End Sub
Private Sub chkCondicion_Click()
    If chkCondicion.Value = 1 Then
        ActivarControl tdbtDiarioCaja, False
        ActivarControl tdbtDiarioCajaEgr, True
        ActivarControl tdbtDiarioCajaIng, True
        
        tdbtDiarioCaja.Text = ""
        tdbtDescripCaja.Text = ""
    Else
        ActivarControl tdbtDiarioCaja, True
        ActivarControl tdbtDiarioCajaEgr, False
        ActivarControl tdbtDiarioCajaIng, False
        
        tdbtDiarioCajaEgr.Text = ""
        tdbtDiarioCajaIng.Text = ""
        tdbtDescripCajaEgr.Text = ""
        tdbtDescripCajaIng.Text = ""
        
    End If
    fValidaVariables
End Sub

Private Sub Grabar(ByVal NMensaje As Integer)
Dim RCImpMa As String
    '-----------------------
    Call GrabaOtrosParamIniciales
    '-----------------------

    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    ' *** Grabando Parametros
    On Local Error GoTo ErrorEjecucion
        
    
    If tdbcBaseImponible.BoundText = "" Then
        Mensajes "Seleccione un tipo de Base Imponible por defecto para el Registro de Compras"
        tdbcBaseImponible.BoundText = ""
        pSetFocus tdbcBaseImponible
        Exit Sub
    End If
    
    If tdbcBaseImponibleVtas.BoundText = "" Then
        Mensajes "Seleccione un tipo de Base Imponible por defecto para el Registro de Ventas"
        tdbcBaseImponibleVtas.BoundText = ""
        pSetFocus tdbcBaseImponibleVtas
        Exit Sub
    End If
    
    sqlSp = "spCNT_CONFIG_LIBROS 'BUSCARREGISTRO','" & gsEmpresa & "','','','','','','','','',0,'','','','','','','','','','','','','','','','" & gsAnio & "',''"
    
    '----------------------
    If optPlan(0).Value = True Then
        gsTipoPlan = 0
    Else
        gsTipoPlan = 1
    End If
    '-------------------
    gsPLE = 1 'IIf(chkPle.Value = 1, chkPle.Value, 0) 'frt_202011
    '----------------------
    Call CargaArregloVar
    
    If fValidaDuplicado(sqlSp) Then
        lArrMnt(0) = "EDITAR"
    Else
        lArrMnt(0) = "INSERTAR"
    End If
    
    'If chkPle.Value = 1 Then 'frt_202011
        frmMDIConta.mnuLibroElec.Enabled = True
    'Else
    '    frmMDIConta.mnuLibroElec.Enabled = False
    'End If
    
    If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCNT_CONFIG_LIBROS", lArrMnt(), True) Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
'    RCImpMa = Trim(tdbImpresora.Text)
'    Dim i As Integer
'    Dim impresora As String
'    impresora = "" + "'?%&$#!¡¿*¨][:´{()/\°|={}"
'    For i = 1 To Len(impresora)
'    RCImpMa = Replace(RCImpMa, Mid(impresora, i, 1), "")
'    Next
'    RCImpMa = "Impresora=" + RCImpMa
'    ExisteFile ("C:\ECBWIN\Rpts Formato Matricial")
'    Guardar_Archivo "C:\ECBWIN\Impresora.ini", ""
'    Abrir_ArchivoBat ("C:\ECBWIN\Impresora.ini")
'    Guardar_Archivo "C:\ECBWIN\Impresora.ini", RCImpMa
    
    
    gsBaseImpDefCom = tdbcBaseImponible.BoundText
    gsBaseImpDefVtas = tdbcBaseImponibleVtas.BoundText
    gsDiarioSimplificado = NE(chkDiarioSimplificado.Value)
    gsNumDigDiarioSimpRep = NE(tdbnDigReporte.Value)
    gintLECompraSimplificado = NE(Me.chkLECompra.Value)
    gintLEVentaSimplificado = NE(Me.chkLEVenta.Value)
    gstrVersionLE = CE(Me.chkVersionLE.Value)
    gsRVIE = CE(Me.chkRVIE.Value)
    ''hlp20230801
    If NMensaje = 1 Then
       Mensajes "Los datos se grabaron con exito...", vbInformation
    End If
     ''hlp20230801
    '----------------------------------------
    Call frmMDIConta.ActivaMenuSegunTipoPlan
    '----------------------------------------

    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub LlenaComboBaseImponible()
    Dim lrsTipo As New ADODB.Recordset
    With lrsTipo
        .CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
        .CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
        .LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
        .Fields.Append "CODIGO", adChar, 3
        .Fields.Append "DESCRIPCION", adVarChar, 50
        .Open
        .AddNew: lrsTipo.Fields("CODIGO") = "   ": .Fields("DESCRIPCION") = "<Seleccione un tipo de Base Imponible>"
        .AddNew: lrsTipo.Fields("CODIGO") = "006": .Fields("DESCRIPCION") = "(A) DEST. A OP.GRAV Y/O EXPORTACION"
        .AddNew: lrsTipo.Fields("CODIGO") = "007": .Fields("DESCRIPCION") = "(B) DEST. A OP.GRAV Y/O EXP. Y NO GRAV."
        .AddNew: lrsTipo.Fields("CODIGO") = "008": .Fields("DESCRIPCION") = "(C) DEST. A OP. NO GRAVADAS"
        .Update
    End With
    
    Set tdbcBaseImponible.RowSource = lrsTipo
    tdbcBaseImponible.ListField = "DESCRIPCION"
    tdbcBaseImponible.BoundColumn = "CODIGO"
End Sub

Private Sub LlenarCombosNivel()
    Dim lrsNivel As New ADODB.Recordset
    With lrsNivel
        .CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
        .CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
        .LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
        .Fields.Append "CODIGO", adChar, 1
        .Fields.Append "DESCRIPCION", adVarChar, 50
        .Open
        .AddNew: .Fields("CODIGO") = " ": .Fields("DESCRIPCION") = "<Seleccione un nivel>"
        .AddNew: .Fields("CODIGO") = "P": .Fields("DESCRIPCION") = "PRIMER NIVEL"
        .AddNew: .Fields("CODIGO") = "S": .Fields("DESCRIPCION") = "SEGUNDO NIVEL"
        .AddNew: .Fields("CODIGO") = "T": .Fields("DESCRIPCION") = "TERCER NIVEL"
        .AddNew: .Fields("CODIGO") = "C": .Fields("DESCRIPCION") = "CUARTO NIVEL"
        .Update
    End With
    
    Set tdbcNivel.RowSource = lrsNivel
    tdbcNivel.ListField = "DESCRIPCION"
    tdbcNivel.BoundColumn = "CODIGO"
End Sub
Sub CargaArregloVar()
    ReDim lArrMnt(34) As Variant

    lArrMnt(1) = gsEmpresa                  ' Empresa
    lArrMnt(2) = tdbtDiarioCompras.Text     ' COMPRAS
    lArrMnt(3) = tdbtDiarioVentas.Text      ' VENTAS
    If chkCondicion.Value = 1 Then
        lArrMnt(4) = ""                     ' CAJA
        lArrMnt(5) = tdbtDiarioCajaIng.Text ' CAJA INGRESOS
        lArrMnt(6) = tdbtDiarioCajaEgr.Text ' CAJA EGRESOS
    Else
        lArrMnt(4) = tdbtDiarioCaja.Text    ' CAJA
        lArrMnt(5) = ""                     ' CAJA INGRESOS
        lArrMnt(6) = ""                     ' CAJA EGRESOS
    End If
    lArrMnt(7) = ""                         ' HONORARIOS
    lArrMnt(8) = ""                         ' tdbtPercepcion
    lArrMnt(9) = ""         ' tdbtRetencion
    lArrMnt(10) = tdbnIGV.Text              ' IGV
    lArrMnt(11) = "A"                       ' Estado
    lArrMnt(12) = gsUsuario                 ' Usuario
    lArrMnt(13) = gsUsuario                 ' Usuario
    lArrMnt(14) = tdbtDiario.Text           ' DIARIO
    lArrMnt(15) = tdbtDiarioDif.Text        ' DIF CAMBIO
    lArrMnt(16) = tdbtDiarioCierre.Text     ' CIERRE
    lArrMnt(17) = tdbcNivel.BoundText       ' NIVEL CENTRO COSTOS
    lArrMnt(18) = tdbnMesCompras.Value      ' #meses permitidos antes de la fehca de ingreso libro compras
    lArrMnt(19) = tdbnNdigitos.Value        ' N. digitos plan de cuentas
    lArrMnt(20) = tdbtDiarioApe.Text
    lArrMnt(21) = tdbcBaseImponible.BoundText
    lArrMnt(22) = gsTipoPlan
    lArrMnt(23) = chkDiarioSimplificado.Value
    lArrMnt(24) = NE(tdbnDigReporte.Value)
    lArrMnt(25) = tdbcBaseImponibleVtas.BoundText
    lArrMnt(26) = gsAnio
    lArrMnt(27) = 1 'chkPle.Value 'frt_202011
    lArrMnt(28) = Me.chkLEVenta.Value
    lArrMnt(29) = Me.chkLECompra.Value
    lArrMnt(30) = CE(Me.txtIdTransferencia.Text)
    lArrMnt(31) = CE(Me.txtIdTransAutomatica.Text)
    lArrMnt(32) = CE(Me.txtIdNIF.Text)
    lArrMnt(33) = Me.chkVersionLE.Value
    lArrMnt(34) = Me.chkRVIE.Value 'frt_rvie
    lsLibroAjusteNIF = CE(Me.txtIdNIF.Text)
    gintLEVentaSimplificado = CE(Me.chkLEVenta.Value)
    gintLECompraSimplificado = CE(Me.chkLECompra.Value)
    lsLibroTransferenciaCancelacion = CE(Me.txtIdTransferencia.Text)
    lsLibroTransCancAutomatico = CE(Me.txtIdTransAutomatica.Text)
    gstrVersionLE = Me.chkVersionLE.Value
    gsRVIE = Me.chkRVIE.Value 'frt_rvie
End Sub

Private Sub Salir()
    Unload Me
End Sub

Private Sub chkRVIE_Click()
  If chkRVIE.Value = 1 Then
     chkVersionLE.Value = 0
  ElseIf chkRVIE.Value = 0 Then
     chkVersionLE.Value = 1
  End If
End Sub

Private Sub chkVersionLE_Click()
  If chkVersionLE.Value = 1 Then
     chkRVIE.Value = 0
  ElseIf chkVersionLE.Value = 0 Then
     chkRVIE.Value = 1
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case vbKeyEscape:
                respuesta = MsgBox("Desea salir de este formulario", vbYesNo + vbQuestion)
                If respuesta = vbYes Then Unload Me
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar (1): KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()
Dim impresora As String
    
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    SSTabParam.Caption = ""
    Centrar_form Me
    
    SSTabParam.TabsPerRow = 1
    SSTabParam.TabVisible(0) = True
    '--------------------------------------------
    ActivarControl tdbtDiarioCaja, True
    ActivarControl tdbtDiarioCajaEgr, False
    ActivarControl tdbtDiarioCajaIng, False
    
    ActivarControl tdbtDescripCaja, False
    ActivarControl tdbtDescripCajaEgr, False
    ActivarControl tdbtDescripCajaIng, False
    ActivarControl tdbtDescripVentas, False
    ActivarControl tdbtDescripCompras, False
    ActivarControl tdbtDescripDiario, False
    ActivarControl tdbtDescripDif, False
    ActivarControl tdbtDescripCierre, False
    ActivarControl tdbtDescripApe, False
    
    '--------------------------------------------
    Call LlenarCombosNivel
    Call LlenaComboBaseImponible
    Call LlenaComboBaseImponibleVentas
    
    
    If SW_ActPLE = "1" Then
        Frame3.Visible = True
'        Frame3.Top = 3600
'        fraDiarioSimplificado.Top = "4250"
        frmMDIConta.mnulin10.Visible = True
        frmMDIConta.mnuLibroElec.Visible = True
    Else
        frmMDIConta.mnuLibroElec.Visible = False
        frmMDIConta.mnulin10.Visible = False
        fraDiarioSimplificado.Top = "4000"
        Frame3.Visible = False
        Frame3.Enabled = False
    End If
    
    
    'RCImpMa = "Impresora=" + RCImpMa
    
'    ExisteFile ("C:\ECBWIN\Rpts Formato Matricial")
'    If Not ExisteArchivo("C:\ECBWIN\Impresora.ini") Then
'    Guardar_Archivo "C:\ECBWIN\Impresora.ini", "Impresora="
'    End If
'    impresora = Trim(Abrir_ArchivoBat("C:\ECBWIN\Impresora.ini"))
'    impresora = Right(impresora, Len(impresora) - 10)
'    Close #1
'    If Len(impresora) = 2 Then tdbImpresora.Text = "" Else tdbImpresora.Text = Trim(impresora)
   
    DoEvents
        
    Call LlenaOtrasConfigIniciales
    ConectarAdvance
     gcnSistemaAdv.Execute "spCn_UpdPeriodoParamInic '" & gsEmpresa & "','" & gsAnio & "'"
    Desconectar
    Call pLlenarVariablesSubDiarios
      
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        
        Call ActivarBotonGrabar(False)
    Else
        Call ActivarBotonGrabar(True)
        
        Call fValidaVariables
    End If
    
    
    'If gsAMBOSPLANDECUENTAS = True Then optPlan(1).Visible = True
End Sub

Private Sub ActivarBotonGrabar(bValor As Boolean)
    tbrOpciones.Buttons(5).Enabled = bValor
End Sub

Private Sub Imprimir()
    Dim matriz(13) As Variant
    Dim Titulo As String
    Titulo = "Reporte de Libro - Tipo de documento"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;LIBRO;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;;True"
    matriz(5) = "@Titulo04;CODIGO;True"
    matriz(6) = "@Titulo05;DESCRIP;True"
    matriz(7) = "@Titulo06;ABREV.;True"
    matriz(8) = "@Titulo07;;True"
    matriz(9) = "@Tipo;LIBRO_TIPODOC;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandarAgrupado.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTabParam
            .Width = Me.Width - 200
            .Height = Me.Height - .Top - 500
            'Frame1.Width = .Width - 200
            'Frame1.Height = .Height - 200
        End With
        
        Call Centrar_Objeto(Frame1, SSTabParam, 0, 200)
       
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTdoc)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
'    CerrarRecordSet lrsDiario
End Sub

Private Sub GrabaNivel()
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    ' *** Eliminando la Cuenta
    Screen.MousePointer = vbHourglass
   
    Dim lArr(17) As Variant
    lArr(0) = "EDITARNIVEL"
    lArr(1) = gsEmpresa
    lArr(2) = Null
    lArr(3) = Null
    lArr(4) = Null
    lArr(5) = Null
    lArr(6) = Null
    lArr(7) = Null
    lArr(8) = Null
    lArr(9) = Null
    lArr(10) = Null
    lArr(11) = "A"
    lArr(12) = gsUsuario
    lArr(13) = gsUsuario
    lArr(14) = Null
    lArr(15) = Null
    lArr(16) = Null
    lArr(17) = tdbcNivel.BoundText ' Left(tdbcNivel.Text, 1)

    
    If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCNT_CONFIG_LIBROS", lArr(), True) Then
        Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
    End If
    
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub GrabaOtrosParamIniciales()
    Call GrabaCondigOPDet("027", tdbnUIT.Value) 'GRABA UIT
    Call GrabaCondigOPDet("053", tdbnIGV.Value) 'GRABA IGV
    Call GrabaCondigOPDet("054", tdbnNdigitos.Value) 'DIGITOS CUENTA
    
    Call GrabaNivel
End Sub


Function fValidaDuplicado(sql) As Boolean
Dim objCon As clsMantoTablas
Dim rs As ADODB.Recordset
Dim arr() As Variant
'Dim sql As String
    arr = Array(sql)
    Set objCon = New clsMantoTablas
    Set rs = objCon.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
    If rs Is Nothing Then
        fValidaDuplicado = False: Exit Function
    End If
    fValidaDuplicado = True
    rs.Close
    Set rs = Nothing
End Function

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
       Case Is = "3" 'GRABA
            Call Grabar(1) ''hlp20230801
       Case Is = "7" 'CANCELAR
            If MsgBox("Desea salir del mantenimiento de Parametros Iniciales", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                Unload Me
            End If
    End Select
End Sub

Private Sub tdbtDiarioApe_Change()
    pTextChange tdbtDiarioApe, tdbtDescripApe
    fValidaVariables

End Sub

Private Sub tdbtDiarioApe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioApe", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiarioCaja_Change()
    pTextChange tdbtDiarioCaja, tdbtDescripCaja
    fValidaVariables
End Sub

Private Sub tdbtDiarioCajaEgr_Change()
    pTextChange tdbtDiarioCajaEgr, tdbtDescripCajaEgr
    fValidaVariables
End Sub

Private Sub tdbtDiarioCajaIng_Change()
    pTextChange tdbtDiarioCajaIng, tdbtDescripCajaIng
    fValidaVariables
End Sub

Private Sub tdbtDiarioCierre_Change()
    pTextChange tdbtDiarioCierre, tdbtDescripCierre
    fValidaVariables

End Sub

Private Sub tdbtDiarioCierre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioCierre", Control, "Libros", Me, gsPeriodo)
    End If

End Sub

Private Sub tdbtDiarioCompras_Change()
    pTextChange tdbtDiarioCompras, tdbtDescripCompras
    fValidaVariables
End Sub

Private Sub tdbtDiarioDif_Change()
    pTextChange tdbtDiarioDif, tdbtDescripDif
    fValidaVariables

End Sub

Private Sub tdbtDiarioDif_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioDif", Control, "Libros", Me, gsPeriodo)
End If
End Sub

Private Sub tdbtDiarioVentas_Change()
    pTextChange tdbtDiarioVentas, tdbtDescripVentas
    fValidaVariables
End Sub

Private Sub tdbtDiario_Change()
    pTextChange tdbtDiario, tdbtDescripDiario
    fValidaVariables
End Sub

Private Sub tdbtDiarioCaja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioCaja", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiarioCajaEgr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioCajaEgr", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiarioCajaIng_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioCajaIng", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiarioCompras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioCompras", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiario", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDiarioVentas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "tdbtDiarioVentas", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
    Select Case Control

        Case "tdbtDiarioVentas"
            tdbtDiarioVentas = Trim(param0)
            tdbtDescripVentas = Trim(param1)
            pSetFocus tdbtDiarioVentas
        Case "tdbtDiarioCompras"
            tdbtDiarioCompras = Trim(param0)
            tdbtDescripCompras = Trim(param1)
            pSetFocus tdbtDiarioCompras

        Case "tdbtDiarioDif"
            tdbtDiarioDif = Trim(param0)
            tdbtDescripDif = Trim(param1)
            pSetFocus tdbtDiarioDif
    
        Case "tdbtDiarioCierre"
            tdbtDiarioCierre = Trim(param0)
            tdbtDescripCierre = Trim(param1)
            pSetFocus tdbtDiarioCierre

        Case "tdbtDiarioCaja"
            tdbtDiarioCaja = Trim(param0)
            tdbtDescripCaja = Trim(param1)
            pSetFocus tdbtDiarioCaja

        Case "tdbtDiarioCajaIng"
            tdbtDiarioCajaIng = Trim(param0)
            tdbtDescripCajaIng = Trim(param1)
            pSetFocus tdbtDiarioCajaIng

        Case "tdbtDiarioCajaEgr"
            tdbtDiarioCajaEgr = Trim(param0)
            tdbtDescripCajaEgr = Trim(param1)
            pSetFocus tdbtDiarioCajaEgr
            
        Case "tdbtDiario"
            tdbtDiario = Trim(param0)
            tdbtDescripDiario = Trim(param1)
            pSetFocus tdbtDiario
            
        Case "tdbtDiarioApe"
            tdbtDiarioApe = Trim(param0)
            tdbtDescripApe = Trim(param1)
            pSetFocus tdbtDiarioApe
        Case "txtIdTransferencia"
            txtIdTransferencia = Trim(param0)
            Me.txtDescripcionTransferencia = Trim(param1)
            pSetFocus txtIdTransferencia
        Case "txtIdTransAutomatica"
            Me.txtIdTransAutomatica = Trim(param0)
            Me.txtDescripTransAutomatico = Trim(param1)
            pSetFocus Me.txtIdTransAutomatica
        Case "txtIdNIF"
            Me.txtIdNIF.Text = Trim(param0)
            Me.txtDescripcionNIF.Text = Trim(param1)
            pSetFocus Me.txtIdNIF
    End Select
    Unload frmBuscador
End Sub

Function fDevuelveDescripcion(sql As String) As String
Dim objCon As clsMantoTablas
Dim rs As ADODB.Recordset
Dim arr() As Variant
arr = Array(sql)
Set objCon = New clsMantoTablas
Set rs = objCon.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
If rs Is Nothing Then
    MsgBox "Código No Existe", vbOKOnly + vbInformation, gsNombreModulo
    fDevuelveDescripcion = ""
    Exit Function
End If
fDevuelveDescripcion = NuloText(rs!Descripcion)
rs.Close
Set rs = Nothing
End Function

Private Sub pTextChange(oTxtCodigo As TDBText, oTxtDescrip As TDBText)
    If Len(oTxtCodigo) < oTxtCodigo.MaxLength Then
        oTxtDescrip.Text = ""
    Else
        pTextLostFocus oTxtCodigo, oTxtDescrip
    End If
    
End Sub

Private Sub pTextLostFocus(oTxtCodigo As TDBText, oTxtDescrip As TDBText)
    If Len(oTxtCodigo) < 2 Then Exit Sub
    sqlSp = "spCn_GrabaLibroOpera 'BUSCARREGISTRO','" & gsEmpresa & "','" & gsAnio & "','" & oTxtCodigo.Text & "'"
    oTxtDescrip = fDevuelveDescripcion(sqlSp)
    
End Sub

Function fValidaVariables() As Boolean
    
    Call ActivarBotonGrabar(True)
    
    '------------------------
    If Len(CE(tdbtDiarioVentas.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    If Len(CE(tdbtDiarioCompras.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    If Len(CE(tdbtDiarioDif.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    If Len(CE(tdbtDiarioCierre.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    If Len(CE(tdbtDiario.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    If Len(CE(tdbtDiarioApe.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    
    If chkCondicion.Value = 1 Then
        If Len(CE(tdbtDiarioCajaIng.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
        If Len(CE(tdbtDiarioCajaEgr.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    Else
        If Len(CE(tdbtDiarioCaja.Text)) < 2 Then Call ActivarBotonGrabar(False): Exit Function
    End If
    
    
    'VALIDA QUE LOS CODIGOS NO SEAN IGUALES
    
    If CE(tdbtDiarioVentas) = CE(tdbtDiarioCompras) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiario) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioCajaIng) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioCajaEgr) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioCaja) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiarioVentas) = CE(tdbtDiario) Then
        Call ActivarBotonGrabar(False)
    End If
    
    If CE(tdbtDiarioCompras) = CE(tdbtDiarioVentas) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiario) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioCajaIng) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioCajaEgr) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioCaja) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiarioCompras) = CE(tdbtDiario) Then
        
        Call ActivarBotonGrabar(False)
    End If
    
    If CE(tdbtDiario) = CE(tdbtDiarioVentas) Or _
        CE(tdbtDiario) = CE(tdbtDiarioCompras) Or _
        CE(tdbtDiario) = CE(tdbtDiarioCajaIng) Or _
        CE(tdbtDiario) = CE(tdbtDiarioCajaEgr) Or _
        CE(tdbtDiario) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiario) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiario) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiario) = CE(tdbtDiarioCaja) Then
        
        Call ActivarBotonGrabar(False)
    End If

    If CE(tdbtDiarioDif) = CE(tdbtDiarioVentas) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioCompras) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioCajaIng) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioCajaEgr) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioCaja) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioCierre) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiarioApe) Or _
       CE(tdbtDiarioDif) = CE(tdbtDiario) Then
        
       Call ActivarBotonGrabar(False)
    End If
    
    If CE(tdbtDiarioCierre) = CE(tdbtDiarioVentas) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioCompras) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioCajaIng) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioCajaEgr) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioCaja) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioDif) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiarioApe) Or _
       CE(tdbtDiarioCierre) = CE(tdbtDiario) Then
        
       Call ActivarBotonGrabar(False)
    End If

    If CE(tdbtDiarioApe) = CE(tdbtDiarioVentas) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioCompras) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioCajaIng) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioCajaEgr) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioCaja) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioDif) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiario) Or _
       CE(tdbtDiarioApe) = CE(tdbtDiarioCierre) Then

       Call ActivarBotonGrabar(False)
    End If

    If (CE(tdbtDiarioCaja) = CE(tdbtDiarioVentas) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioCompras) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioCajaIng) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioCajaEgr) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiarioCaja) = CE(tdbtDiario)) And chkCondicion.Value = vbUnchecked Then
        
        Call ActivarBotonGrabar(False)
    End If

    If (CE(tdbtDiarioCajaIng) = CE(tdbtDiarioVentas) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioCompras) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioCaja) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioCajaEgr) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiarioCajaIng) = CE(tdbtDiario)) And chkCondicion.Value = vbChecked Then
        
        Call ActivarBotonGrabar(False)
    End If

    If (CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioVentas) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioCompras) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioCajaIng) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioCaja) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioDif) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioCierre) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiarioApe) Or _
        CE(tdbtDiarioCajaEgr) = CE(tdbtDiario)) And chkCondicion.Value = vbChecked Then
        
        Call ActivarBotonGrabar(False)
    End If
    
    
    If CE(tdbtDescripVentas) = "" Or _
       CE(tdbtDescripCompras) = "" Or _
       CE(tdbtDescripDiario) = "" Or _
       CE(tdbtDescripDif) = "" Or _
       CE(tdbtDescripCierre) = "" Or _
       CE(tdbtDescripApe) = "" Or _
       CE(tdbtDescripDiario) = "" Then
        Call ActivarBotonGrabar(False)
    End If
    
    If CE(tdbtDescripCaja) = "" And chkCondicion.Value = vbUnchecked Then
        Call ActivarBotonGrabar(False)
    End If
    
    If (CE(tdbtDescripCajaEgr) = "" Or CE(tdbtDescripCajaIng) = "") And chkCondicion.Value = vbChecked Then
        Call ActivarBotonGrabar(False)
    End If
        
End Function
Private Sub LlenaComboBaseImponibleVentas()

    Dim lrsTipo As New ADODB.Recordset
    With lrsTipo
        .CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
        .CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
        .LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
        .Fields.Append "CODIGO", adChar, 3
        .Fields.Append "DESCRIPCION", adVarChar, 50
        .Open
        .AddNew: lrsTipo.Fields("CODIGO") = "   ": .Fields("DESCRIPCION") = "<Seleccione un tipo de Base Imponible de Ventas>"
        .AddNew: lrsTipo.Fields("CODIGO") = "002": .Fields("DESCRIPCION") = "GRAVABLE VENTAS"
        .AddNew: lrsTipo.Fields("CODIGO") = "021": .Fields("DESCRIPCION") = "EXPORTACIONES"
        .AddNew: lrsTipo.Fields("CODIGO") = "998": .Fields("DESCRIPCION") = "EXONERADA"
        .AddNew: lrsTipo.Fields("CODIGO") = "999": .Fields("DESCRIPCION") = "INAFECTO"
        .Update
    End With
    
    Set tdbcBaseImponibleVtas.RowSource = lrsTipo
    tdbcBaseImponibleVtas.ListField = "DESCRIPCION"
    tdbcBaseImponibleVtas.BoundColumn = "CODIGO"
    
End Sub

Private Sub txtIdNIF_Change()
    pTextChange Me.txtIdNIF, Me.txtDescripcionNIF
    fValidaVariables
End Sub

Private Sub txtIdNIF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "txtIdNIF", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub txtIdTransAutomatica_Change()
    pTextChange Me.txtIdTransAutomatica, Me.txtDescripTransAutomatico
    fValidaVariables
End Sub

Private Sub txtIdTransAutomatica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "txtIdTransAutomatica", Control, "Libros", Me, gsPeriodo)
    End If
End Sub

Private Sub txtIdTransferencia_Change()
    pTextChange txtIdTransferencia, Me.txtDescripcionTransferencia
    fValidaVariables
End Sub

Private Sub txtIdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "txtIdTransferencia", Control, "Libros", Me, gsPeriodo)
    End If
End Sub
