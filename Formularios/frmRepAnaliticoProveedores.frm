VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepAnaliticoProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Analisis por Entidades"
   ClientHeight    =   4920
   ClientLeft      =   1245
   ClientTop       =   2115
   ClientWidth     =   9240
   Icon            =   "frmRepAnaliticoProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   9240
   Begin VB.Frame fraTodo 
      Height          =   4845
      Left            =   45
      TabIndex        =   18
      Top             =   0
      Width           =   9105
      Begin VB.Frame Frame4 
         Height          =   4110
         Left            =   135
         TabIndex        =   19
         Top             =   135
         Width           =   8850
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   3765
            Left            =   5220
            TabIndex        =   30
            Top             =   270
            Width           =   3495
            Begin VB.CheckBox chkFVcto 
               Caption         =   "Filtrar por Fecha Vencimiento"
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
               Left            =   140
               TabIndex        =   35
               Top             =   2160
               Width           =   3255
            End
            Begin VB.Frame fraOrden 
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   0
               TabIndex        =   31
               Top             =   2400
               Width           =   3420
               Begin VB.OptionButton optOrden3 
                  Caption         =   "Por Fecha Vencimiento"
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
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1070
                  Width           =   3180
               End
               Begin VB.OptionButton optOrden2 
                  Caption         =   "Por Fecha"
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
                  Left            =   120
                  TabIndex        =   15
                  Top             =   765
                  Width           =   3180
               End
               Begin VB.OptionButton optOrden1 
                  Caption         =   "Por Documento"
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
                  Left            =   120
                  TabIndex        =   14
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   3165
               End
               Begin VB.Label lblORDENADOPOR 
                  AutoSize        =   -1  'True
                  Caption         =   "ORDENADO POR"
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
                  Index           =   0
                  Left            =   135
                  TabIndex        =   32
                  Top             =   180
                  Width           =   1395
               End
            End
            Begin VB.OptionButton optDetalleDoc 
               Caption         =   "Detalle de Documentos Pendientes"
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
               Left            =   120
               TabIndex        =   11
               Top             =   1215
               Visible         =   0   'False
               Width           =   3480
            End
            Begin VB.OptionButton optLetXPagar 
               Caption         =   "Letras Pendientes por Pagar"
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
               Left            =   120
               TabIndex        =   13
               Top             =   1800
               Width           =   3210
            End
            Begin VB.OptionButton optLetXCobrar 
               Caption         =   "Letras Pendientes por Cobrar"
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
               Left            =   120
               TabIndex        =   12
               Top             =   1530
               Width           =   3210
            End
            Begin VB.OptionButton optResumen 
               Caption         =   "Resumen por Entidad y Periodo"
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
               Left            =   120
               TabIndex        =   10
               Top             =   900
               Width           =   3030
            End
            Begin VB.OptionButton optPendientes 
               Caption         =   "Documentos Pendientes de Pago"
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
               Left            =   120
               TabIndex        =   9
               Top             =   630
               Width           =   3210
            End
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos los Documentos"
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
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Value           =   -1  'True
               Width           =   2490
            End
            Begin VB.Label lblTIPODE 
               AutoSize        =   -1  'True
               Caption         =   "TIPO DE REPORTE"
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
               Index           =   1
               Left            =   120
               TabIndex        =   33
               Top             =   45
               Width           =   1530
            End
         End
         Begin VB.Frame FraRANGODE 
            Caption         =   "RANGO DE FECHAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   90
            TabIndex        =   27
            Top             =   195
            Width           =   4995
            Begin VB.CheckBox chkTodosAnios 
               Caption         =   "Buscar en años anteriores"
               Height          =   330
               Left            =   2565
               TabIndex        =   2
               Top             =   675
               Width           =   2175
            End
            Begin TDBDate6Ctl.TDBDate dtpDesde 
               Height          =   300
               Left            =   930
               TabIndex        =   0
               Tag             =   "enabled"
               Top             =   330
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   529
               Calendar        =   "frmRepAnaliticoProveedores.frx":0ECA
               Caption         =   "frmRepAnaliticoProveedores.frx":0FCC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":1030
               Keys            =   "frmRepAnaliticoProveedores.frx":104E
               Spin            =   "frmRepAnaliticoProveedores.frx":10BA
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
               Text            =   "03/08/2004"
               ValidateMode    =   0
               ValueVT         =   2010185735
               Value           =   38202
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate dtpHasta 
               Height          =   300
               Left            =   3450
               TabIndex        =   1
               Tag             =   "enabled"
               Top             =   315
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   529
               Calendar        =   "frmRepAnaliticoProveedores.frx":10E2
               Caption         =   "frmRepAnaliticoProveedores.frx":11E4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":1248
               Keys            =   "frmRepAnaliticoProveedores.frx":1266
               Spin            =   "frmRepAnaliticoProveedores.frx":12D2
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
               Text            =   "03/08/2004"
               ValidateMode    =   0
               ValueVT         =   2010185735
               Value           =   38202
               CenturyMode     =   0
            End
            Begin VB.Label lblDesde 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   210
               TabIndex        =   29
               Top             =   390
               Width           =   555
            End
            Begin VB.Label lblHasta 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   2565
               TabIndex        =   28
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   90
            TabIndex        =   22
            Top             =   1560
            Width           =   4980
            Begin VB.CheckBox chkTodos 
               Caption         =   "RANGO DE CUENTAS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   135
               TabIndex        =   3
               Top             =   0
               Width           =   2265
            End
            Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
               Height          =   315
               Left            =   1125
               TabIndex        =   4
               Tag             =   "_"
               Top             =   360
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "frmRepAnaliticoProveedores.frx":12FA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":1366
               Key             =   "frmRepAnaliticoProveedores.frx":1384
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
            Begin TDBText6Ctl.TDBText tdbtDescripcionDesde 
               Height          =   315
               Left            =   2445
               TabIndex        =   23
               Tag             =   "_"
               Top             =   360
               Width           =   2400
               _Version        =   65536
               _ExtentX        =   4233
               _ExtentY        =   556
               Caption         =   "frmRepAnaliticoProveedores.frx":13D6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":1442
               Key             =   "frmRepAnaliticoProveedores.frx":1460
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   1
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
               Left            =   1125
               TabIndex        =   5
               Tag             =   "_"
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "frmRepAnaliticoProveedores.frx":14B2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":151E
               Key             =   "frmRepAnaliticoProveedores.frx":153C
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
            Begin TDBText6Ctl.TDBText tdbtDescripcionHasta 
               Height          =   315
               Left            =   2445
               TabIndex        =   24
               Tag             =   "_"
               Top             =   720
               Width           =   2400
               _Version        =   65536
               _ExtentX        =   4233
               _ExtentY        =   556
               Caption         =   "frmRepAnaliticoProveedores.frx":158E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":15FA
               Key             =   "frmRepAnaliticoProveedores.frx":1618
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   1
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
            Begin VB.Label lblCtaHasta 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cta. Hasta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   165
               TabIndex        =   26
               Top             =   765
               Width           =   870
            End
            Begin VB.Label lblCtaDesde 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cta. Desde"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   165
               TabIndex        =   25
               Top             =   405
               Width           =   930
            End
         End
         Begin VB.Frame FraENTIDADES 
            Caption         =   " ENTIDADES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   90
            TabIndex        =   20
            Top             =   3120
            Width           =   4965
            Begin TrueOleDBList70.TDBCombo tdbcTipoEntidad 
               Height          =   300
               Left            =   1140
               TabIndex        =   6
               Tag             =   "_"
               Top             =   270
               Width           =   2460
               _ExtentX        =   4339
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
               _PropDict       =   $"frmRepAnaliticoProveedores.frx":166A
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
            Begin TDBText6Ctl.TDBText tdbtCodigo 
               Height          =   315
               Left            =   3660
               TabIndex        =   7
               Top             =   270
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   556
               Caption         =   "frmRepAnaliticoProveedores.frx":16F1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRepAnaliticoProveedores.frx":175D
               Key             =   "frmRepAnaliticoProveedores.frx":177B
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
            Begin VB.Label lblEntidad 
               AutoSize        =   -1  'True
               Caption         =   "Entidad"
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
               Left            =   180
               TabIndex        =   21
               Top             =   315
               Width           =   630
            End
         End
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   4680
         TabIndex        =   17
         Top             =   4290
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepAnaliticoProveedores.frx":17CD
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   2760
         TabIndex        =   16
         Top             =   4290
         Width           =   1665
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmRepAnaliticoProveedores.frx":1D67
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
      TabIndex        =   34
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepAnaliticoProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Sub chkTodos_Click()
    If chkTodos.Value = vbUnchecked Then
        tdbtCuentaDesde = ""
        tdbtCuentaHasta = ""
        tdbtDescripcionDesde = ""
        tdbtDescripcionHasta = ""
        
        tdbtCuentaDesde.ReadOnly = True
        tdbtCuentaHasta.ReadOnly = True
        tdbtCuentaDesde.BackColor = gsColorDesactivado
        tdbtCuentaHasta.BackColor = gsColorDesactivado
        tdbtDescripcionDesde.BackColor = gsColorDesactivado
        tdbtDescripcionHasta.BackColor = gsColorDesactivado
        
    Else
        tdbtCuentaDesde.ReadOnly = False
        tdbtCuentaHasta.ReadOnly = False
        tdbtCuentaDesde.BackColor = gsColorActivado
        tdbtCuentaHasta.BackColor = gsColorActivado
        tdbtDescripcionDesde.BackColor = gsColorDesactivado
        tdbtDescripcionHasta.BackColor = gsColorDesactivado
    End If
End Sub

Private Sub chkTodosAnios_Click()
    
    If chkTodosAnios.Value = vbChecked Then
       dtpDesde.MinDate = "01/01/1900"
       
     Else
       dtpDesde.MinDate = "01/01/" & gsAnio
       dtpDesde = "01/01/" & gsAnio
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    Dim Tipo As String
    Dim matriz_fecha(14) As Variant
    Dim sqlver  As String, valorDato As String
    Dim TipoBusq As String
    
    If chkTodos.Value = vbChecked And (tdbtCuentaDesde = "" Or tdbtCuentaHasta = "") Then
       Mensajes "Ingrese las cuentas inicial y final", vbOKOnly + vbInformation
       Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    cmdImprimir.Enabled = False
    DoEvents
    Tipo = ""
    matriz_fecha(0) = "@Tipo;" & Tipo & ";True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz_fecha(3) = "@desde;" & dtpDesde.Text & ";True"
    matriz_fecha(4) = "@hasta;" & dtpHasta.Text & ";True"
    matriz_fecha(5) = "@Pla_cCuentaContable;" & tdbtCuentaDesde & ";True"
    matriz_fecha(6) = "@cTipoEntidad;" & tdbcTipoEntidad.BoundText & ";True"
    matriz_fecha(7) = "@cCodEntidad;" & tdbtCodigo.Text & ";True"
    matriz_fecha(8) = "@TodoAnio;" & chkTodosAnios.Value & ";True"
    matriz_fecha(9) = "@usuario;" & gsUsuario & ";True"
    matriz_fecha(10) = "@TD;;True"
    matriz_fecha(11) = "@Pla_cCuentaContable2;" & tdbtCuentaHasta & ";True"
    
    matriz_fecha(12) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(13) = "@RUC;" & "RUC : " & gsRUC & ";True"

    TipoBusq = "N"
    matriz_fecha(14) = "@TipoBusqueda;" & TipoBusq & ";True"
    If chkFVcto.Value Then
        TipoBusq = "V"
        matriz_fecha(14) = "@TipoBusqueda;" & TipoBusq & ";True"
    End If

    Dim formulas(2) As Variant

    If optResumen.Value = True Then
        If Me.optOrden1.Value = True Then
            formulas(0) = "grupo02 = {spCn_RptAnalisisDocumentos;1.Ent_nRuc}"
            formulas(1) = "grupo03 = {spCn_RptAnalisisDocumentos;1.Per_cPeriodo}"
        Else
            formulas(0) = "grupo02 = {spCn_RptAnalisisDocumentos;1.Per_cPeriodo}"
            formulas(1) = "grupo03 = {spCn_RptAnalisisDocumentos;1.Ent_nRuc}"
        End If
    Else
        If Me.optOrden1.Value = True Then
            formulas(0) = "orden02 = {spCn_RptAnalisisDocumentos;1.Asd_cSerieDoc}+{spCn_RptAnalisisDocumentos;1.Asd_cNumDoc}"
            formulas(1) = "orden01 = {spCn_RptAnalisisDocumentos;1.Asd_dFecDoc}"
            formulas(2) = "orden03 = {spCn_RptAnalisisDocumentos;1.Asd_dFecVen}"
        ElseIf Me.optOrden2.Value = True Then
            formulas(0) = "orden02 = {spCn_RptAnalisisDocumentos;1.Asd_dFecDoc}"
            formulas(1) = "orden01 = {spCn_RptAnalisisDocumentos;1.Asd_cSerieDoc}+{spCn_RptAnalisisDocumentos;1.Asd_cNumDoc}"
            formulas(2) = "orden03 = {spCn_RptAnalisisDocumentos;1.Asd_dFecVen}"
        Else
            formulas(0) = "orden02 = {spCn_RptAnalisisDocumentos;1.Asd_dFecVen}"
            formulas(1) = "orden01 = {spCn_RptAnalisisDocumentos;1.Asd_cSerieDoc}+{spCn_RptAnalisisDocumentos;1.Asd_cNumDoc}"
            formulas(2) = "orden03 = {spCn_RptAnalisisDocumentos;1.Asd_dFecDoc}"
        End If
    End If
    '----------------------
    If optTodos.Value = True Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptAnalisisDocumentos.rpt", crptToWindow, "Reporte de Documentos", "", matriz_fecha(), formulas()
    End If
    '----------------------
    If optPendientes.Value = True Then
        matriz_fecha(0) = "@Tipo;SALDOS;True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptAnalisisSaldosDocumentos.rpt", crptToWindow, "Reporte de Documentos Pendientes", "", matriz_fecha(), formulas()
    End If
    '----------------------
    If optResumen.Value = True Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDocumentosResumen.rpt", crptToWindow, "Resumen Entidad y Periodo", "", matriz_fecha(), formulas()
    End If
    '----------------------
    'If optResumenDoc.Value = True Then AbreReporteParam gsDSN, Me, rutaReportes & "RptAnalisisSaldosDocumentos.rpt", crptToWindow, "Resumen Entidad y Documento", "", matriz_fecha(), Formulas()
    '----------------------
    If optLetXCobrar.Value = True Then
    
        sqlver = "SELECT Cod_cValorParam FROM CND_CONFIG_OPERA " & _
                 "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and  " & _
                 "cop_ccodigo='023'"
                 
        valorDato = ExtraeDescripcion(sqlver)
        matriz_fecha(0) = "@Tipo;LETRA_COBRAR;True"
        matriz_fecha(10) = "@TD;" & valorDato & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDocumentosPendientesxCobrar.rpt", crptToWindow, "Reporte de Letras por Cobrar Pendientes", "", matriz_fecha(), formulas()
    End If
    '----------------------
    If optLetXPagar.Value = True Then
    
        sqlver = "SELECT Cod_cValorParam FROM CND_CONFIG_OPERA " & _
                 "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and  " & _
                 "cop_ccodigo='022'"
                 
        matriz_fecha(0) = "@Tipo;LETRA_PAGAR;True"
        valorDato = ExtraeDescripcion(sqlver)
        matriz_fecha(10) = "@TD;" & valorDato & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDocumentosPendientesxPagar.rpt", crptToWindow, "Reporte de Letras por Pagar Pendientes", "", matriz_fecha(), formulas()
    End If
    
    Screen.MousePointer = vbNormal
    cmdImprimir.Enabled = True
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
        
    Dim VarMes As String
    Dim sqlcombos As String
    
    Call Centrar_form(Me)
    
    If gsPeriodo = "00" Then
        VarMes = "01/01/" + gsAnio
        dtpHasta = UltimoDiaMes("01", gsAnio)
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        VarMes = "01/12/" + gsAnio
        dtpHasta = UltimoDiaMes("12", gsAnio)
    Else
        VarMes = "01/" & gsPeriodo & "/" + gsAnio
        dtpHasta = UltimoDiaMes(gsPeriodo, gsAnio)
    End If
    
    dtpDesde = VarMes
        
    ' *** Llenando el tipo de Entidad
    sqlcombos = "SELECT Ten_cTipoEntidad, Ten_cNombreEntidad From CNT_ENTIDAD "
    sqlcombos = sqlcombos + "WHERE Emp_cCodigo = '" & gsEmpresa & "' ORDER BY Ten_cNombreEntidad"
    LlenarComboAddItem tdbcTipoEntidad, sqlcombos, True
    
    dtpDesde.MinDate = "01/01/" & gsAnio
    dtpDesde.MaxDate = "31/12/" & gsAnio
    dtpHasta.MinDate = "01/01/" & gsAnio
    dtpHasta.MaxDate = "31/12/" & gsAnio
    
    chkTodos.Value = vbUnchecked
    chkTodos_Click
    tdbcTipoEntidad_ItemChange
    Call optTodos_Click
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fratodo, Me)
        Call CentrarTitulo(lblTitulo, fratodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepAnaliticoProveedores = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub optDetalleDoc_Click()
    fraOrden.Visible = True
    chkFVcto.Value = False
    chkFVcto.Enabled = False
    optOrden3.Visible = False
End Sub

Private Sub optLetXCobrar_Click()
    optOrden1.Caption = "Por Documento"
    optOrden2.Caption = "Por Fecha"
    optOrden1.Value = True
    fraOrden.Visible = True
    optOrden3.Visible = True
    If optLetXCobrar.Value Then chkFVcto.Enabled = True Else chkFVcto.Enabled = False: chkFVcto.Value = False
End Sub

Private Sub optLetXPagar_Click()
    optOrden1.Caption = "Por Documento"
    optOrden2.Caption = "Por Fecha"
    optOrden1.Value = True
    fraOrden.Visible = True
    optOrden3.Visible = True
    If optLetXPagar.Value Then chkFVcto.Enabled = True Else chkFVcto.Enabled = False: chkFVcto.Value = False
End Sub

Private Sub optOrden1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If

End Sub

Private Sub optOrden2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If
End Sub


Private Sub optOrden3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If
End Sub

Private Sub optPendientes_Click()
    optOrden1.Caption = "Por Documento"
    optOrden2.Caption = "Por Fecha"
    optOrden1.Value = True
    fraOrden.Visible = False
    chkFVcto.Value = False
    chkFVcto.Enabled = False
    optOrden3.Visible = False
End Sub

Private Sub optPendientes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optOrden1.Value Then
        pSetFocus optOrden1
    Else
        pSetFocus optOrden2
    End If
End If
End Sub

Private Sub optResumen_Click()
    optOrden1.Caption = "Por Entidad / Periodo"
    optOrden2.Caption = "Por Periodo / Entidad"
    optOrden1.Value = True
    fraOrden.Visible = True
    chkFVcto.Value = False
    chkFVcto.Enabled = False
    optOrden3.Visible = False
End Sub

Private Sub optResumen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optOrden1.Value Then
        pSetFocus optOrden1
    Else
        pSetFocus optOrden2
    End If
End If
End Sub

Private Sub optTodos_Click()
    optOrden1.Caption = "Por Documento"
    optOrden2.Caption = "Por Fecha"
    optOrden1.Value = True
    fraOrden.Visible = False
    chkFVcto.Value = False
    chkFVcto.Enabled = False
    optOrden3.Visible = False
End Sub

Private Sub optTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optOrden1.Value Then
        pSetFocus optOrden1
    Else
        pSetFocus optOrden2
    End If
End If
End Sub

Private Sub tdbcTipoEntidad_ItemChange()
    If tdbcTipoEntidad.BoundText = "" Then
        tdbtCodigo = ""
        ActivarControl tdbtCodigo, False
    Else
        ActivarControl tdbtCodigo, True
    End If
End Sub

Private Sub tdbcTipoEntidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtCodigo
    If tdbtCodigo.Enabled = True Then
       pSetFocus tdbtCodigo
    Else
       pSetFocus cmdImprimir
    End If
End If
End Sub

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(tdbcTipoEntidad.BoundText) = "" And tdbcTipoEntidad.BoundText <> "" Then
       Mensajes "Debe ingresar el Tipo de Entidad, verificar...", vbInformation
       pSetFocus tdbtCodigo
       Exit Sub
    End If

    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCodigo.Name, Control, "Entidad", Me, gsPeriodo, tdbcTipoEntidad.BoundText)
    
End Sub

Private Sub tdbtCodigo_LostFocus()
    Dim valorDato As String, sqlver As String
        
    If tdbtCodigo <> "" And Me.Enabled = True Then
        tdbtCodigo.Text = Right("00000" & CE(tdbtCodigo.Text), 5)
        If Trim(tdbcTipoEntidad.BoundText) = "" Then
           Mensajes "Debe ingresar el Tipo de Entidad, verificar...", vbInformation
           pSetFocus tdbcTipoEntidad
           Exit Sub
        End If
            
        sqlver = "SELECT ENT_CPERSONA From CNM_ENTIDAD "
        sqlver = sqlver + "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND  ent_cEstado = 'A'"
        sqlver = sqlver + "AND Ent_cCodEntidad = '" & tdbtCodigo.Text & "' "
        sqlver = sqlver + "AND Ten_CTipoEntidad = '" & Trim(tdbcTipoEntidad.BoundText) & "' "
        
        valorDato = ExtraeDescripcion(sqlver)
        If valorDato = "" Then
           Mensajes "Codigo de entidad no existe, verificar...", vbInformation
           tdbtCodigo.Text = ""
           pSetFocus tdbtCodigo
           
        End If
    End If
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Call LlamaBuscar(frmBuscador, Me.tdbtCuentaDesde.Name, Control, "Cuentas", Me, gsPeriodo, tdbtCuentaDesde.Text)
    End If
    If KeyCode = 13 Then
        pSetFocus tdbtCuentaHasta
    End If
    
End Sub

Private Sub tdbtCuentaDesde_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(CE(tdbtCuentaDesde.Text)) = 0 Then
        tdbtDescripcionDesde.Text = ""
    End If
End Sub

Private Sub tdbtCuentaDesde_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCuentaDesde <> "" And Me.Enabled = True Then
        tdbtDescripcionDesde = ExisteCtaNoTitulo(tdbtCuentaDesde, "")
        If tdbtDescripcionDesde = "" Then pSetFocus tdbtCuentaDesde
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
    Case "tdbtCuentaDesde" ' *** Caso Desde
        tdbtCuentaDesde = Trim(param0)
        tdbtDescripcionDesde = Trim(param1)
        Unload frmBuscador
        pSetFocus tdbtCuentaDesde
    Case "tdbtCodigo" '     *** Caso Codigp
        tdbtCodigo = Trim(param0)
        Unload frmBuscador
        pSetFocus tdbtCodigo
    Case "tdbtCuentaHasta" ' *** Caso Desde
        tdbtCuentaHasta = Trim(param0)
        tdbtDescripcionHasta = Trim(param1)
        Unload frmBuscador
        pSetFocus tdbtCuentaDesde
        
    End Select
End Sub

Private Sub tdbtCuentaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Call LlamaBuscar(frmBuscador, Me.tdbtCuentaHasta.Name, Control, "Cuentas", Me, gsPeriodo, tdbtCuentaHasta.Text)
    End If
    If KeyCode = 13 Then
        pSetFocus tdbcTipoEntidad
    End If
    
End Sub

Private Sub tdbtCuentaHasta_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(CE(tdbtCuentaHasta.Text)) = 0 Then
        tdbtDescripcionHasta.Text = ""
    End If
End Sub

Private Sub tdbtCuentaHasta_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCuentaHasta <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta = ExisteCtaNoTitulo(tdbtCuentaHasta, "")
        If tdbtDescripcionHasta = "" Then pSetFocus tdbtCuentaHasta
    End If

End Sub
