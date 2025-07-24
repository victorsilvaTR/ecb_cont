VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmManSubDiarioTDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros Iniciales"
   ClientHeight    =   6090
   ClientLeft      =   3405
   ClientTop       =   3600
   ClientWidth     =   11655
   Icon            =   "frmManSubDiarioTDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   11655
   Begin TabDlg.SSTab SSTabParam 
      Height          =   5595
      Left            =   60
      TabIndex        =   0
      Top             =   405
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9869
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
         Height          =   5370
         Left            =   90
         TabIndex        =   1
         Top             =   45
         Width           =   11205
         Begin VB.Frame Frame3 
            Caption         =   "Libros Electrónicos"
            Height          =   600
            Left            =   7800
            TabIndex        =   53
            Top             =   3600
            Width           =   3075
            Begin VB.CheckBox chkPle 
               Caption         =   "   Habilitar  PLE"
               Height          =   195
               Left            =   600
               TabIndex        =   54
               Top             =   270
               Width           =   1695
            End
         End
         Begin VB.Frame fraDiarioSimplificado 
            Caption         =   "Diario Formato Simplificado"
            Height          =   945
            Left            =   7800
            TabIndex        =   47
            Top             =   4260
            Width           =   3075
            Begin VB.CheckBox chkDiarioSimplificado 
               Alignment       =   1  'Right Justify
               Caption         =   "Habilitar"
               Height          =   195
               Left            =   330
               TabIndex        =   48
               Top             =   285
               Width           =   2250
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnDigReporte 
               Height          =   285
               Left            =   1935
               TabIndex        =   49
               Top             =   525
               Width           =   660
               _Version        =   65536
               _ExtentX        =   1164
               _ExtentY        =   503
               Calculator      =   "frmManSubDiarioTDoc.frx":0EE6
               Caption         =   "frmManSubDiarioTDoc.frx":0F06
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManSubDiarioTDoc.frx":0F6A
               Keys            =   "frmManSubDiarioTDoc.frx":0F88
               Spin            =   "frmManSubDiarioTDoc.frx":0FE0
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
               ValueVT         =   2293761
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
            Top             =   2595
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
            Left            =   -45
            TabIndex        =   2
            Top             =   2430
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
            Caption         =   "frmManSubDiarioTDoc.frx":1008
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1074
            Key             =   "frmManSubDiarioTDoc.frx":1092
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
            Caption         =   "frmManSubDiarioTDoc.frx":10D4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1140
            Key             =   "frmManSubDiarioTDoc.frx":115E
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
            Caption         =   "frmManSubDiarioTDoc.frx":11A0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":120C
            Key             =   "frmManSubDiarioTDoc.frx":122A
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
            Caption         =   "frmManSubDiarioTDoc.frx":126C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":12D8
            Key             =   "frmManSubDiarioTDoc.frx":12F6
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
            Caption         =   "frmManSubDiarioTDoc.frx":1338
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":13A4
            Key             =   "frmManSubDiarioTDoc.frx":13C2
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
            Caption         =   "frmManSubDiarioTDoc.frx":1404
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1470
            Key             =   "frmManSubDiarioTDoc.frx":148E
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
            Caption         =   "frmManSubDiarioTDoc.frx":14D0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":153C
            Key             =   "frmManSubDiarioTDoc.frx":155A
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
            Caption         =   "frmManSubDiarioTDoc.frx":159C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1608
            Key             =   "frmManSubDiarioTDoc.frx":1626
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
            Caption         =   "frmManSubDiarioTDoc.frx":1668
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":16D4
            Key             =   "frmManSubDiarioTDoc.frx":16F2
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
            Caption         =   "frmManSubDiarioTDoc.frx":1734
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":17A0
            Key             =   "frmManSubDiarioTDoc.frx":17BE
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
            Caption         =   "frmManSubDiarioTDoc.frx":1800
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":186C
            Key             =   "frmManSubDiarioTDoc.frx":188A
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
            Caption         =   "frmManSubDiarioTDoc.frx":18CC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1938
            Key             =   "frmManSubDiarioTDoc.frx":1956
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
            Left            =   6570
            TabIndex        =   20
            Top             =   4110
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":1998
            Caption         =   "frmManSubDiarioTDoc.frx":19B8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1A1C
            Keys            =   "frmManSubDiarioTDoc.frx":1A3A
            Spin            =   "frmManSubDiarioTDoc.frx":1A92
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
            ValueVT         =   1769473
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
            Caption         =   "frmManSubDiarioTDoc.frx":1ABA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1B26
            Key             =   "frmManSubDiarioTDoc.frx":1B44
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
            Caption         =   "frmManSubDiarioTDoc.frx":1B86
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1BF2
            Key             =   "frmManSubDiarioTDoc.frx":1C10
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
            Caption         =   "frmManSubDiarioTDoc.frx":1C52
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1CBE
            Key             =   "frmManSubDiarioTDoc.frx":1CDC
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
            Caption         =   "frmManSubDiarioTDoc.frx":1D1E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1D8A
            Key             =   "frmManSubDiarioTDoc.frx":1DA8
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
            Caption         =   "frmManSubDiarioTDoc.frx":1DEA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1E56
            Key             =   "frmManSubDiarioTDoc.frx":1E74
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
            Caption         =   "frmManSubDiarioTDoc.frx":1EB6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":1F22
            Key             =   "frmManSubDiarioTDoc.frx":1F40
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
            Left            =   6570
            TabIndex        =   21
            Top             =   4695
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":1F82
            Caption         =   "frmManSubDiarioTDoc.frx":1FA2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":2006
            Keys            =   "frmManSubDiarioTDoc.frx":2024
            Spin            =   "frmManSubDiarioTDoc.frx":207C
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
            ValueVT         =   1769473
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnMesCompras 
            Height          =   285
            Left            =   3945
            TabIndex        =   18
            Top             =   4035
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":20A4
            Caption         =   "frmManSubDiarioTDoc.frx":20C4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":2128
            Keys            =   "frmManSubDiarioTDoc.frx":2146
            Spin            =   "frmManSubDiarioTDoc.frx":219E
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
            ValueVT         =   1507329
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnNdigitos 
            Height          =   285
            Left            =   3945
            TabIndex        =   19
            Top             =   4680
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   503
            Calculator      =   "frmManSubDiarioTDoc.frx":21C6
            Caption         =   "frmManSubDiarioTDoc.frx":21E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManSubDiarioTDoc.frx":224A
            Keys            =   "frmManSubDiarioTDoc.frx":2268
            Spin            =   "frmManSubDiarioTDoc.frx":22C0
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
            ValueVT         =   1769473
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSDataListLib.DataCombo tdbcBaseImponible 
            Height          =   300
            Left            =   3960
            TabIndex        =   17
            Top             =   2970
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
            Left            =   3960
            TabIndex        =   16
            Top             =   2610
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
            Left            =   3960
            TabIndex        =   52
            Top             =   3360
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
         Begin VB.Label Label4 
            Caption         =   "Base Imponible x defecto Reg. Ventas"
            Height          =   270
            Index           =   7
            Left            =   225
            TabIndex        =   51
            Top             =   3375
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
            Left            =   5670
            TabIndex        =   40
            Top             =   4125
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
            Left            =   240
            TabIndex        =   32
            Top             =   2610
            Width           =   2970
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "U.I.T."
            Height          =   195
            Index           =   2
            Left            =   5700
            TabIndex        =   31
            Top             =   4740
            Width           =   405
         End
         Begin VB.Label Label4 
            Caption         =   "Meses ant. permitidos en ingreso docs. en Libro de Compras"
            Height          =   420
            Index           =   3
            Left            =   210
            TabIndex        =   30
            Top             =   4005
            Width           =   3585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "N° Digitos Cuenta Detalle, Plan cuentas"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   29
            Top             =   4695
            Width           =   2820
         End
         Begin VB.Label Label4 
            Caption         =   "Base Imponible x defecto Reg. Compras"
            Height          =   270
            Index           =   5
            Left            =   240
            TabIndex        =   28
            Top             =   2940
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
            Picture         =   "frmManSubDiarioTDoc.frx":22E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2442
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":259C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":26F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2850
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":29AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":2C5E
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
            Picture         =   "frmManSubDiarioTDoc.frx":2DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":3352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":38EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":3E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":4420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":49BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":4F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSubDiarioTDoc.frx":54EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   43
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

    tdbnMesCompras.Value = NE(rs!Cfl_cMesCompras)
 
    If CE(rs!Cfl_cBaseDefCom) <> "006" And CE(rs!Cfl_cBaseDefCom) <> "007" And CE(rs!Cfl_cBaseDefCom) <> "008" Then
        tdbcBaseImponible.BoundText = ""
    Else
        tdbcBaseImponible.BoundText = CE(rs!Cfl_cBaseDefCom)
    End If
    
    tdbcBaseImponibleVtas.BoundText = CE(rs!Cfl_cBaseDefVtas)
    
    tdbcNivel.BoundText = CE(rs!Cfl_cNivelCC)
    chkDiarioSimplificado.Value = vbUnchecked
    
    chkPle.Value = IIf(IsNull(rs!Cfl_cPle) = True, 0, IIf(rs!Cfl_cPle = 0, 0, 1))

    On Error Resume Next
    gsTipoPlan = NE(rs!Cfl_cTipoPlan)
    optPlan(gsTipoPlan).Value = True
    tdbnDigReporte.Value = NE(rs!Cfl_nDigDiarioRep)
    
    chkDiarioSimplificado.Value = NE(rs!Cfl_cDiarioSimplificado)
    Call CerrarRecordSet(rs)
    Set rs = Nothing
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

Private Sub Grabar()
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
    gsPLE = IIf(chkPle.Value = 1, chkPle.Value, 0)
    '----------------------
    Call CargaArregloVar
    
    If fValidaDuplicado(sqlSp) Then
        lArrMnt(0) = "EDITAR"
    Else
        lArrMnt(0) = "INSERTAR"
    End If
    
    If chkPle.Value = 1 Then
        frmMDIConta.mnuLibroElec.Enabled = True
    Else
        frmMDIConta.mnuLibroElec.Enabled = False
    End If
    
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
    
    Mensajes "Los datos se grabaron con exito...", vbInformation
    
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
    ReDim lArrMnt(27) As Variant
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
    
    'PGBV - 11012013
'    If chkPle.Value = True Then
        lArrMnt(27) = chkPle.Value
'    Else
'        lArrMnt(27) = 0
'    End If
    
End Sub

Private Sub Salir()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case vbKeyEscape:
                respuesta = MsgBox("Desea salir de este formulario", vbYesNo + vbQuestion)
                If respuesta = vbYes Then Unload Me
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar: KeyCode = 0
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
        Frame3.Top = 3600
        fraDiarioSimplificado.Top = "4250"
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
    Dim periodo As String
    Select Case Button.Index
       Case Is = "3" 'GRABA
                      
            Call Grabar
            
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
