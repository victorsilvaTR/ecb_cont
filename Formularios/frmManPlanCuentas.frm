VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManPlanCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Plan de Cuentas"
   ClientHeight    =   7980
   ClientLeft      =   2250
   ClientTop       =   2100
   ClientWidth     =   10875
   Icon            =   "frmManPlanCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmManPlanCuentas.frx":0ECA
   ScaleHeight     =   7980
   ScaleWidth      =   10875
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
            Picture         =   "frmManPlanCuentas.frx":120C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":1366
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":161A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":1774
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":18CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":1A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":1B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":1CDC
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
            Picture         =   "frmManPlanCuentas.frx":1E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":23D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":296A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":2F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":349E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":3A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":3FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":456C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":4B06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstPerfiles 
      Height          =   7365
      Left            =   135
      TabIndex        =   23
      Top             =   465
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   12991
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta Plan de Cuentas"
      TabPicture(0)   =   "frmManPlanCuentas.frx":50A0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Plan de Cuentas"
      TabPicture(1)   =   "frmManPlanCuentas.frx":50BC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cuentas por Dif Cambio"
      TabPicture(2)   =   "frmManPlanCuentas.frx":50D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDifCambio"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraDifCambio 
         Height          =   4695
         Left            =   -74550
         TabIndex        =   39
         Top             =   660
         Width           =   8475
         Begin VB.Frame Frame8 
            Caption         =   "Redondeo"
            Height          =   1275
            Left            =   225
            TabIndex        =   65
            Top             =   2745
            Width           =   8115
            Begin TDBText6Ctl.TDBText tdbtCodRedondeoG 
               Height          =   315
               Left            =   1995
               TabIndex        =   66
               Top             =   345
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":50F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":5160
               Key             =   "frmManPlanCuentas.frx":517E
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtDescRedondeoG 
               Height          =   315
               Left            =   3750
               TabIndex        =   67
               Top             =   345
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":51D0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":523C
               Key             =   "frmManPlanCuentas.frx":525A
               BackColor       =   16777152
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtCodRedondeoP 
               Height          =   315
               Left            =   1995
               TabIndex        =   68
               Top             =   825
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":52AC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":5318
               Key             =   "frmManPlanCuentas.frx":5336
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtDescRedondeoP 
               Height          =   315
               Left            =   3765
               TabIndex        =   69
               Top             =   825
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":5388
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":53F4
               Key             =   "frmManPlanCuentas.frx":5412
               BackColor       =   16777152
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta de Pérdida"
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
               Left            =   75
               TabIndex        =   71
               Top             =   840
               Width           =   1545
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta de Ganancia"
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
               Left            =   75
               TabIndex        =   70
               Top             =   390
               Width           =   1695
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Diferencia de Cambio"
            Height          =   1230
            Left            =   225
            TabIndex        =   43
            Top             =   1485
            Width           =   8115
            Begin TDBText6Ctl.TDBText tdbtCodGanancia 
               Height          =   315
               Left            =   1995
               TabIndex        =   44
               Top             =   360
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":5464
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":54D0
               Key             =   "frmManPlanCuentas.frx":54EE
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtCodPerdida 
               Height          =   315
               Left            =   1980
               TabIndex        =   46
               Top             =   765
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":5540
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":55AC
               Key             =   "frmManPlanCuentas.frx":55CA
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtDescGanancia 
               Height          =   315
               Left            =   3750
               TabIndex        =   45
               Top             =   360
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":561C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":5688
               Key             =   "frmManPlanCuentas.frx":56A6
               BackColor       =   16777152
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin TDBText6Ctl.TDBText tdbtDescPerdida 
               Height          =   315
               Left            =   3735
               TabIndex        =   47
               Top             =   765
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               Caption         =   "frmManPlanCuentas.frx":56F8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManPlanCuentas.frx":5764
               Key             =   "frmManPlanCuentas.frx":5782
               BackColor       =   16777152
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta de Ganancia"
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
               Left            =   75
               TabIndex        =   49
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta de Pérdida"
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
               Index           =   9
               Left            =   75
               TabIndex        =   48
               Top             =   825
               Width           =   1545
            End
         End
         Begin MSForms.CommandButton cmdActualizar 
            Height          =   435
            Left            =   4500
            TabIndex        =   58
            Top             =   4140
            Width           =   1665
            Caption         =   "  Grabar"
            PicturePosition =   327683
            Size            =   "2937;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEditar 
            Height          =   435
            Left            =   2250
            TabIndex        =   57
            Top             =   4140
            Width           =   1665
            Caption         =   " Editar datos"
            PicturePosition =   327683
            Size            =   "2937;767"
            Picture         =   "frmManPlanCuentas.frx":57D4
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   360
            Picture         =   "frmManPlanCuentas.frx":5D6E
            Stretch         =   -1  'True
            Top             =   495
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Estas pueden ser editadas, desde el asiento. Si se requiere utilizar mas de una cuenta para el Tipo de Cambio."
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
            Left            =   2295
            TabIndex        =   41
            Top             =   945
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Las siguientes son las cuentas que se utilizaran por defecto, en la generación Automatica de Asientos por Diferencia de Cambio."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2310
            TabIndex        =   40
            Top             =   405
            Width           =   5415
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6900
         Left            =   195
         TabIndex        =   25
         Top             =   360
         Width           =   10125
         Begin VB.CheckBox chkTitulo 
            Alignment       =   1  'Right Justify
            Caption         =   "Titulo"
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
            Left            =   5790
            TabIndex        =   5
            Tag             =   "_"
            Top             =   405
            Width           =   1695
         End
         Begin TabDlg.SSTab sstParamatros 
            Height          =   5520
            Left            =   120
            TabIndex        =   28
            Top             =   1215
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   9737
            _Version        =   393216
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Parametros"
            TabPicture(0)   =   "frmManPlanCuentas.frx":6078
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraDatos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Destino"
            TabPicture(1)   =   "frmManPlanCuentas.frx":6094
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraDestinoCuenta"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Configuración"
            TabPicture(2)   =   "frmManPlanCuentas.frx":60B0
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fraTexto"
            Tab(2).Control(1)=   "fraConfig"
            Tab(2).ControlCount=   2
            Begin VB.Frame fraConfig 
               Height          =   4905
               Left            =   -74850
               TabIndex        =   74
               Top             =   345
               Width           =   8880
               Begin VB.CheckBox chkCostoProduccion 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Costo Produccion"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   171
                  Top             =   4440
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.CheckBox chkVariacionProduccion 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Variacion Produccion"
                  Height          =   195
                  Left            =   4680
                  TabIndex        =   169
                  Top             =   3960
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.CheckBox chkCuentaCostoVenta 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Cuenta Costo de Venta"
                  Height          =   195
                  Left            =   1440
                  TabIndex        =   167
                  Top             =   3960
                  Width           =   2415
               End
               Begin VB.Frame fraNoTitulos 
                  BorderStyle     =   0  'None
                  Height          =   4140
                  Left            =   480
                  TabIndex        =   87
                  Top             =   360
                  Width           =   7830
                  Begin VB.ComboBox CmbRPI 
                     Height          =   315
                     Left            =   6750
                     Style           =   2  'Dropdown List
                     TabIndex        =   164
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   735
                  End
                  Begin VB.CheckBox ChkConsPDT601 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Considerar para PDT 601 - PLAME."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   3585
                     TabIndex        =   156
                     Top             =   2235
                     Width           =   3870
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de Letras por Pagar"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   10
                     Left            =   3780
                     TabIndex        =   124
                     Top             =   3825
                     Visible         =   0   'False
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de Letras por Cobrar"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   9
                     Left            =   3780
                     TabIndex        =   123
                     Top             =   3630
                     Visible         =   0   'False
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta Pagar/Cobrar con Base Imp."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   8
                     Left            =   45
                     TabIndex        =   94
                     Top             =   660
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de Cuarta Quinta Especial"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   7
                     Left            =   45
                     TabIndex        =   120
                     Top             =   3195
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de Quinta Especial"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   6
                     Left            =   45
                     TabIndex        =   119
                     Top             =   2955
                     Width           =   2940
                  End
                  Begin TabDlg.SSTab SSTabCierre 
                     Height          =   600
                     Left            =   6960
                     TabIndex        =   97
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1845
                     _ExtentX        =   3254
                     _ExtentY        =   1058
                     _Version        =   393216
                     TabOrientation  =   1
                     Style           =   1
                     TabHeight       =   520
                     TabCaption(0)   =   " Cargas Imputables "
                     TabPicture(0)   =   "frmManPlanCuentas.frx":60CC
                     Tab(0).ControlEnabled=   -1  'True
                     Tab(0).Control(0)=   "fraNoTituloCta7"
                     Tab(0).Control(0).Enabled=   0   'False
                     Tab(0).ControlCount=   1
                     TabCaption(1)   =   " Result Ejercicio "
                     TabPicture(1)   =   "frmManPlanCuentas.frx":60E8
                     Tab(1).ControlEnabled=   0   'False
                     Tab(1).Control(0)=   "fraNoTituloCta456"
                     Tab(1).ControlCount=   1
                     TabCaption(2)   =   "Result. Explotación"
                     TabPicture(2)   =   "frmManPlanCuentas.frx":6104
                     Tab(2).ControlEnabled=   0   'False
                     Tab(2).ControlCount=   0
                     Begin VB.Frame fraNoTituloCta456 
                        BackColor       =   &H00FFFFFF&
                        BorderStyle     =   0  'None
                        Height          =   2625
                        Left            =   -74955
                        TabIndex        =   103
                        Top             =   45
                        Width           =   4515
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de Reserva Legal"
                           Height          =   285
                           Index           =   5
                           Left            =   360
                           TabIndex        =   109
                           Top             =   1665
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de variación de existencias"
                           Height          =   285
                           Index           =   4
                           Left            =   360
                           TabIndex        =   108
                           Top             =   1350
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de Tributos por Pagar"
                           Height          =   285
                           Index           =   3
                           Left            =   360
                           TabIndex        =   107
                           Top             =   1035
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de Remuneraciones y Particip. por Pagar"
                           Height          =   285
                           Index           =   2
                           Left            =   360
                           TabIndex        =   106
                           Top             =   720
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de Perdida"
                           Height          =   285
                           Index           =   1
                           Left            =   360
                           TabIndex        =   105
                           Top             =   405
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreVarias 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Cuenta de Utilidad"
                           Height          =   285
                           Index           =   0
                           Left            =   360
                           TabIndex        =   104
                           Top             =   90
                           Width           =   3840
                        End
                     End
                     Begin VB.Frame fraNoTituloCta7 
                        BackColor       =   &H00FFFFFF&
                        BorderStyle     =   0  'None
                        Height          =   2685
                        Left            =   420
                        TabIndex        =   98
                        Top             =   330
                        Width           =   4545
                        Begin VB.CheckBox chkCierreCargas 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Gastos Finacieros"
                           Height          =   285
                           Index           =   3
                           Left            =   360
                           TabIndex        =   102
                           Top             =   1035
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreCargas 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Gastos Administrativos"
                           Height          =   285
                           Index           =   2
                           Left            =   360
                           TabIndex        =   101
                           Top             =   720
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreCargas 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Gasto de Ventas"
                           Height          =   285
                           Index           =   1
                           Left            =   360
                           TabIndex        =   100
                           Top             =   405
                           Width           =   3840
                        End
                        Begin VB.CheckBox chkCierreCargas 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Costo de Servicios"
                           Height          =   285
                           Index           =   0
                           Left            =   360
                           TabIndex        =   99
                           Top             =   90
                           Width           =   3840
                        End
                     End
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cta. de Bonif. y Transf. Gratuita"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   5
                     Left            =   15
                     TabIndex        =   93
                     Top             =   2415
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Reg. Compras Cta. Reintegro"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   2
                     Left            =   -15
                     MaskColor       =   &H00C0FFFF&
                     TabIndex        =   91
                     Top             =   1290
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de Exportaciones"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   4
                     Left            =   30
                     TabIndex        =   92
                     Top             =   2055
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Cuenta de I.S.C."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   1
                     Left            =   45
                     TabIndex        =   90
                     Top             =   405
                     Width           =   2940
                  End
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Reg. Compras Cta. Honorarios"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   0
                     Left            =   4680
                     TabIndex        =   89
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   3825
                  End
                  Begin TDBText6Ctl.TDBText tdbtPla_cCuenta39 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   162
                     Tag             =   "_"
                     Top             =   1320
                     Visible         =   0   'False
                     Width           =   1395
                     _Version        =   65536
                     _ExtentX        =   2461
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":6120
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":618C
                     Key             =   "frmManPlanCuentas.frx":61AA
                     BackColor       =   16514043
                     EditMode        =   0
                     ForeColor       =   8388608
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
                  Begin VB.CheckBox chkNoTit 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Reg. Compras Otros"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   3
                     Left            =   3810
                     TabIndex        =   95
                     Top             =   2970
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin TDBText6Ctl.TDBText tdbtPla_cCuenta39Nombre 
                     Height          =   315
                     Left            =   5640
                     TabIndex        =   163
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   4935
                     _Version        =   65536
                     _ExtentX        =   8705
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":61FC
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":6268
                     Key             =   "frmManPlanCuentas.frx":6286
                     BackColor       =   14737632
                     EditMode        =   0
                     ForeColor       =   8388608
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
                  Begin VB.Frame fraNoTituloCta8 
                     BorderStyle     =   0  'None
                     Height          =   2595
                     Left            =   3255
                     TabIndex        =   110
                     Top             =   390
                     Width           =   4635
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Margen Comercial"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   0
                        Left            =   360
                        TabIndex        =   118
                        Top             =   360
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Resultado del Ejercicio"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   6
                        Left            =   360
                        TabIndex        =   112
                        Top             =   90
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Excedente Bruto de Explotación"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   2
                        Left            =   360
                        TabIndex        =   116
                        Top             =   870
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Valor Agregado"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   1
                        Left            =   360
                        TabIndex        =   117
                        Top             =   600
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Resultado de Explotación"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   3
                        Left            =   360
                        TabIndex        =   115
                        Top             =   1125
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Resultado antes de Participaciones y Impuestos"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   4
                        Left            =   360
                        TabIndex        =   114
                        Top             =   1410
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000004&
                        Caption         =   "Distribución Legal de Renta"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   5
                        Left            =   360
                        TabIndex        =   113
                        Top             =   1680
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCta8 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Impuesto a la Renta"
                        Height          =   285
                        Index           =   7
                        Left            =   360
                        TabIndex        =   111
                        Top             =   1950
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.CheckBox chkCierreCtaCnfImp 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000000&
                        Caption         =   "(I.G.V) - Impuesto a Operaciones"
                        ForeColor       =   &H80000008&
                        Height          =   285
                        Left            =   360
                        TabIndex        =   155
                        Top             =   855
                        Visible         =   0   'False
                        Width           =   3840
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Regimen Pensionario de Independientes"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   360
                        TabIndex        =   165
                        Top             =   2280
                        Visible         =   0   'False
                        Width           =   3375
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Utilizar como Cuenta de Impuesto"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   285
                        Index           =   10
                        Left            =   360
                        TabIndex        =   154
                        Top             =   540
                        Visible         =   0   'False
                        Width           =   3450
                     End
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Cuenta 39 Asociada"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   11
                     Left            =   3600
                     TabIndex        =   161
                     Top             =   2805
                     Visible         =   0   'False
                     Width           =   1980
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Registro de Compras"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   9
                     Left            =   45
                     TabIndex        =   152
                     Top             =   1035
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Registro de Ventas"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   5
                     Left            =   45
                     TabIndex        =   127
                     Top             =   1785
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     BackColor       =   &H80000004&
                     Caption         =   "Cuentas de Asientos Automáticos"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   4
                     Left            =   3795
                     TabIndex        =   122
                     Top             =   3375
                     Visible         =   0   'False
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Retrib. Inc. F) Art. 34 LIR"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   3
                     Left            =   45
                     TabIndex        =   121
                     Top             =   2715
                     Width           =   2940
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H80000010&
                     BorderWidth     =   2
                     X1              =   3180
                     X2              =   3180
                     Y1              =   90
                     Y2              =   4200
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Cuentas de Cierre de Ejercicio"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   2
                     Left            =   3555
                     TabIndex        =   96
                     Top             =   105
                     Width           =   4260
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Registro de Compras/ Ventas"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   1
                     Left            =   45
                     TabIndex        =   88
                     Top             =   90
                     Width           =   2940
                  End
               End
               Begin VB.Frame fraTitulos 
                  BorderStyle     =   0  'None
                  Height          =   3465
                  Left            =   1320
                  TabIndex        =   75
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   5865
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Cta. Aumento/Reducción del Rep. Valores"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   420
                     Index           =   9
                     Left            =   3240
                     TabIndex        =   85
                     Top             =   1800
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     Caption         =   "Cuenta por Pagar Honorarios"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   8
                     Left            =   3195
                     TabIndex        =   84
                     Top             =   2295
                     Visible         =   0   'False
                     Width           =   2355
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Cuenta por Pagar"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   7
                     Left            =   3240
                     TabIndex        =   83
                     Top             =   1215
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Cuenta por Cobrar"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   6
                     Left            =   3240
                     TabIndex        =   82
                     Top             =   855
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Adq. Grav. Destinadas a Op. No Grav. ( C )"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   420
                     Index           =   4
                     Left            =   90
                     TabIndex        =   81
                     Top             =   2790
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Adq. Grav. Destinadas a Op. Grav. y/o Exportac. y a Op. No Grav. ( B )"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   600
                     Index           =   3
                     Left            =   90
                     TabIndex        =   80
                     Top             =   2115
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Adq. Grav. Destinadas a Op. Grav. y/o Exportac. ( A )"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   420
                     Index           =   2
                     Left            =   90
                     TabIndex        =   79
                     Top             =   1620
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000000&
                     Caption         =   "Base Imponible Ventas"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   1
                     Left            =   90
                     TabIndex        =   77
                     Top             =   855
                     Width           =   2400
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     Caption         =   "Base Imponible Compras"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   0
                     Left            =   3240
                     TabIndex        =   76
                     Top             =   2970
                     Visible         =   0   'False
                     Width           =   2175
                  End
                  Begin VB.CheckBox chkTit 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     Caption         =   "Cuenta de Igv Automatico"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   285
                     Index           =   5
                     Left            =   3240
                     TabIndex        =   78
                     Top             =   2700
                     Visible         =   0   'False
                     Width           =   2220
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Control de Cuentas"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   8
                     Left            =   3240
                     TabIndex        =   151
                     Top             =   540
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Reporte de Reg. de Ventas"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   7
                     Left            =   90
                     TabIndex        =   150
                     Top             =   540
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Reporte de Reg. de Compras"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   6
                     Left            =   90
                     TabIndex        =   149
                     Top             =   1260
                     Width           =   2940
                  End
                  Begin VB.Label Label13 
                     Alignment       =   2  'Center
                     Caption         =   "Solo para cuentas de Titulos"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   0
                     Left            =   90
                     TabIndex        =   86
                     Top             =   90
                     Width           =   5550
                  End
               End
               Begin VB.Label lblCostoProduccion 
                  Caption         =   "Costo Produccion"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   170
                  Top             =   4200
                  Width           =   2175
               End
               Begin VB.Label lblVariacionProduccion 
                  Caption         =   "Variacion Produccion"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   168
                  Top             =   3720
                  Width           =   2175
               End
               Begin VB.Label lblCuentaCostoVenta 
                  Caption         =   "Cuenta Costo Venta"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   166
                  Top             =   3720
                  Width           =   2535
               End
            End
            Begin VB.Frame fraDestinoCuenta 
               Height          =   3585
               Left            =   -74055
               TabIndex        =   32
               Top             =   810
               Width           =   7905
               Begin TrueOleDBList70.TDBList tdblDestinoAux 
                  Height          =   1695
                  Left            =   360
                  TabIndex        =   21
                  Top             =   1800
                  Width           =   7110
                  _ExtentX        =   12541
                  _ExtentY        =   2990
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "Destino Debe"
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "Destino Haber"
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "Nombre Cuenta"
                  Columns(2).DataField=   "Nombre Cuenta"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "%"
                  Columns(3).DataField=   "%"
                  Columns(3).NumberFormat=   "Standard"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   4
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=4"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2117"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2037"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
                  Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(13)=   "Column(2).Width=6271"
                  Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6191"
                  Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
                  Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(19)=   "Column(3).Width=688"
                  Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=609"
                  Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
                  Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   1
                  BorderStyle     =   1
                  MatchEntry      =   0
                  RightToLeft     =   0   'False
                  MatchCompare    =   -6
                  MatchCol        =   0
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   5
                  MultiSelect     =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   1
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  ExposeCellMode  =   0
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  LayoutUrl       =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DataView        =   0
                  GroupByCaption  =   "Drag a column header here to group by that column"
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=248,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
                  _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&"
                  _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
                  _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
                  _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
                  _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
                  _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
                  _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
                  _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
                  _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
                  _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
                  _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
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
               Begin TDBNumber6Ctl.TDBNumber tdbnPorc 
                  Height          =   315
                  Left            =   2475
                  TabIndex        =   20
                  Top             =   1335
                  Width           =   1260
                  _Version        =   65536
                  _ExtentX        =   2222
                  _ExtentY        =   556
                  Calculator      =   "frmManPlanCuentas.frx":62D8
                  Caption         =   "frmManPlanCuentas.frx":62F8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmManPlanCuentas.frx":6364
                  Keys            =   "frmManPlanCuentas.frx":6382
                  Spin            =   "frmManPlanCuentas.frx":63DA
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   1
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "##0.00 %"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   8388608
                  Format          =   "##0.00 %"
                  HighlightText   =   1
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   100
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   0
                  Separator       =   ","
                  ShowContextMenu =   1
                  ValueVT         =   5
                  Value           =   100
                  MaxValueVT      =   1802698757
                  MinValueVT      =   1769209861
               End
               Begin VB.ComboBox cmbTipo 
                  BackColor       =   &H00FBFBFB&
                  ForeColor       =   &H00800000&
                  Height          =   315
                  ItemData        =   "frmManPlanCuentas.frx":6402
                  Left            =   1050
                  List            =   "frmManPlanCuentas.frx":640C
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   1335
                  Width           =   1395
               End
               Begin TDBText6Ctl.TDBText tdbtCtaDestino 
                  Height          =   315
                  Left            =   1050
                  TabIndex        =   18
                  Tag             =   "_"
                  Top             =   885
                  Width           =   1395
                  _Version        =   65536
                  _ExtentX        =   2461
                  _ExtentY        =   556
                  Caption         =   "frmManPlanCuentas.frx":641D
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmManPlanCuentas.frx":6489
                  Key             =   "frmManPlanCuentas.frx":64A7
                  BackColor       =   16514043
                  EditMode        =   0
                  ForeColor       =   8388608
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
                  Format          =   "aA"
                  FormatMode      =   1
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
               Begin TDBText6Ctl.TDBText tdbtNombreDestino 
                  Height          =   315
                  Left            =   2475
                  TabIndex        =   22
                  Top             =   885
                  Width           =   4935
                  _Version        =   65536
                  _ExtentX        =   8705
                  _ExtentY        =   556
                  Caption         =   "frmManPlanCuentas.frx":64F9
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmManPlanCuentas.frx":6565
                  Key             =   "frmManPlanCuentas.frx":6583
                  BackColor       =   14737632
                  EditMode        =   0
                  ForeColor       =   8388608
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
               Begin TrueOleDBList70.TDBCombo tdbcMes 
                  Height          =   300
                  Left            =   1050
                  TabIndex        =   17
                  Top             =   435
                  Width           =   2580
                  _ExtentX        =   4551
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
                  _PropDict       =   $"frmManPlanCuentas.frx":65D5
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFBFBFB&,.fgcolor=&H800000&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               Begin MSForms.CommandButton cmdEliminarDestino 
                  Height          =   435
                  Left            =   5850
                  TabIndex        =   73
                  Top             =   1275
                  Width           =   1545
                  Caption         =   " Eliminar Cuenta"
                  PicturePosition =   327683
                  Size            =   "2725;767"
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   3
               End
               Begin MSForms.CommandButton cmdInsertar 
                  Height          =   435
                  Left            =   4230
                  TabIndex        =   72
                  Top             =   1275
                  Width           =   1545
                  Caption         =   " Insertar Cuenta"
                  PicturePosition =   327683
                  Size            =   "2725;767"
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   3
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Mes :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   210
                  TabIndex        =   42
                  Top             =   450
                  Width           =   750
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tipo :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   3
                  Left            =   225
                  TabIndex        =   34
                  Top             =   1395
                  Width           =   750
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Cuenta :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   2
                  Left            =   225
                  TabIndex        =   33
                  Top             =   930
                  Width           =   750
               End
            End
            Begin VB.Frame fraDatos 
               Height          =   5010
               Left            =   315
               TabIndex        =   29
               Top             =   375
               Width           =   9495
               Begin VB.Frame Frame12 
                  Caption         =   "Cuentas de Resultado"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Left            =   240
                  TabIndex        =   140
                  Top             =   1620
                  Width           =   3375
                  Begin VB.CheckBox chkFun 
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
                     Left            =   1875
                     TabIndex        =   142
                     Top             =   315
                     Width           =   240
                  End
                  Begin VB.CheckBox chkNat 
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
                     Left            =   1875
                     TabIndex        =   141
                     Top             =   615
                     Width           =   255
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Función"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   2
                     Left            =   405
                     TabIndex        =   144
                     Top             =   315
                     Width           =   630
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Naturaleza"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   3
                     Left            =   405
                     TabIndex        =   143
                     Top             =   600
                     Width           =   840
                  End
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Documentos de Referencia y Otros"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1485
                  Left            =   240
                  TabIndex        =   135
                  Top             =   2880
                  Width           =   3375
                  Begin VB.CheckBox ChkNCND 
                     Alignment       =   1  'Right Justify
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
                     Left            =   1890
                     TabIndex        =   139
                     Top             =   315
                     Width           =   200
                  End
                  Begin VB.CheckBox chkDetraccion 
                     Alignment       =   1  'Right Justify
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
                     Left            =   1890
                     TabIndex        =   138
                     Top             =   585
                     Width           =   200
                  End
                  Begin VB.CheckBox chkRetencion 
                     Alignment       =   1  'Right Justify
                     Caption         =   "X"
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
                     Left            =   1890
                     TabIndex        =   137
                     Top             =   855
                     Width           =   200
                  End
                  Begin VB.CheckBox chkPercepcion 
                     Alignment       =   1  'Right Justify
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
                     Left            =   1890
                     TabIndex        =   136
                     Top             =   1125
                     Width           =   200
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Percepción"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   10
                     Left            =   450
                     TabIndex        =   148
                     Top             =   1125
                     Width           =   900
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Retención"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   9
                     Left            =   450
                     TabIndex        =   147
                     Top             =   855
                     Width           =   825
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Detracción"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   8
                     Left            =   450
                     TabIndex        =   146
                     Top             =   585
                     Width           =   870
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Doc. Referencia"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   7
                     Left            =   450
                     TabIndex        =   145
                     Top             =   315
                     Width           =   1290
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "Centro de Costos y Presup."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1125
                  Left            =   3825
                  TabIndex        =   128
                  Top             =   1620
                  Width           =   4725
                  Begin VB.CheckBox chkCentroCosto 
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
                     Left            =   1935
                     TabIndex        =   129
                     Tag             =   "_"
                     Top             =   315
                     Width           =   285
                  End
                  Begin TrueOleDBList70.TDBCombo tdbcPatrimomio 
                     Height          =   300
                     Left            =   1920
                     TabIndex        =   132
                     Tag             =   "_"
                     Top             =   600
                     Width           =   2310
                     _ExtentX        =   4075
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
                     Columns.Count   =   2
                     Splits(0)._UserFlags=   0
                     Splits(0).ExtendRightColumn=   -1  'True
                     Splits(0).AllowRowSizing=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
                     Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                     Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
                     Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                     Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
                     Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
                     Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
                     RowDividerColor =   13160660
                     RowSubDividerColor=   13160660
                     AddItemSeparator=   ";"
                     _PropDict       =   $"frmManPlanCuentas.frx":665C
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=88,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFBFBFB&,.fgcolor=&H800000&"
                     _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
                  Begin VB.Label Label11 
                     Caption         =   "Presupuesto"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   360
                     TabIndex        =   131
                     Top             =   630
                     Width           =   1485
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Centro de costo"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   4
                     Left            =   360
                     TabIndex        =   130
                     Top             =   315
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame9 
                  Caption         =   "Cuentas de Balance"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1245
                  Left            =   240
                  TabIndex        =   56
                  Top             =   285
                  Width           =   3375
                  Begin VB.OptionButton OptPas 
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
                     Left            =   1845
                     TabIndex        =   8
                     Top             =   645
                     Width           =   270
                  End
                  Begin VB.OptionButton OptAct 
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
                     Left            =   1845
                     TabIndex        =   7
                     Top             =   315
                     Value           =   -1  'True
                     Width           =   270
                  End
                  Begin VB.PictureBox Picture1 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   780
                     Left            =   90
                     ScaleHeight     =   780
                     ScaleWidth      =   2895
                     TabIndex        =   59
                     Top             =   225
                     Width           =   2895
                     Begin VB.Label Label12 
                        AutoSize        =   -1  'True
                        BackColor       =   &H00FFFFFF&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Pasivo/Patrim."
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   1
                        Left            =   405
                        TabIndex        =   62
                        Top             =   450
                        Width           =   1140
                     End
                     Begin VB.Label Label12 
                        AutoSize        =   -1  'True
                        BackColor       =   &H00FFFFFF&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Activo"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   0
                        Left            =   405
                        TabIndex        =   61
                        Top             =   90
                        Width           =   510
                     End
                  End
               End
               Begin VB.Frame Frame11 
                  Caption         =   "Campos necesarios para controlar un Documento"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1245
                  Left            =   3825
                  TabIndex        =   54
                  Top             =   285
                  Width           =   4725
                  Begin VB.CheckBox chkProvision 
                     Caption         =   " "
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
                     Left            =   1950
                     TabIndex        =   9
                     Tag             =   "_"
                     Top             =   240
                     Width           =   210
                  End
                  Begin VB.CheckBox chkDocumento 
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
                     Left            =   1950
                     TabIndex        =   10
                     Tag             =   "_"
                     Top             =   495
                     Width           =   255
                  End
                  Begin VB.PictureBox Picture2 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   960
                     Left            =   90
                     ScaleHeight     =   960
                     ScaleWidth      =   4200
                     TabIndex        =   60
                     Top             =   225
                     Width           =   4200
                     Begin TrueOleDBList70.TDBCombo tdbcEntidad 
                        Height          =   300
                        Left            =   1845
                        TabIndex        =   134
                        Tag             =   "_"
                        Top             =   630
                        Width           =   2310
                        _ExtentX        =   4075
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
                        _PropDict       =   $"frmManPlanCuentas.frx":66E3
                        _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                        _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                        _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                        _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                        _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                        _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                        _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFBFBFB&,.fgcolor=&H800000&"
                        _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Entidad"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   13
                        Left            =   240
                        TabIndex        =   133
                        Top             =   645
                        Width           =   1500
                     End
                     Begin VB.Label Label12 
                        AutoSize        =   -1  'True
                        BackColor       =   &H00FFFFFF&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Documento"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   6
                        Left            =   240
                        TabIndex        =   64
                        Top             =   315
                        Width           =   960
                     End
                     Begin VB.Label Label12 
                        AutoSize        =   -1  'True
                        BackColor       =   &H00FFFFFF&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Provisión"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   5
                        Left            =   240
                        TabIndex        =   63
                        Top             =   30
                        Width           =   705
                     End
                  End
               End
               Begin VB.Frame Frame10 
                  Caption         =   " Configuración  Reportes  EEFF"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1785
                  Left            =   3825
                  TabIndex        =   50
                  Top             =   2880
                  Width           =   4725
                  Begin TDBText6Ctl.TDBText tdbtResFunc 
                     Height          =   315
                     Left            =   2595
                     TabIndex        =   14
                     Top             =   645
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":676A
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":67D6
                     Key             =   "frmManPlanCuentas.frx":67F4
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   8388608
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
                     Format          =   "9"
                     FormatMode      =   0
                     AutoConvert     =   -1
                     ErrorBeep       =   0
                     MaxLength       =   4
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
                  Begin TDBText6Ctl.TDBText tdbtResNatu 
                     Height          =   315
                     Left            =   2595
                     TabIndex        =   15
                     Top             =   1035
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":6846
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":68B2
                     Key             =   "frmManPlanCuentas.frx":68D0
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   8388608
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
                     AllowSpace      =   0
                     Format          =   "9"
                     FormatMode      =   0
                     AutoConvert     =   -1
                     ErrorBeep       =   0
                     MaxLength       =   4
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
                  Begin TDBText6Ctl.TDBText tdbtBalance 
                     Height          =   315
                     Left            =   2595
                     TabIndex        =   12
                     Top             =   270
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":6922
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":698E
                     Key             =   "frmManPlanCuentas.frx":69AC
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   8388608
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
                     Format          =   "9"
                     FormatMode      =   0
                     AutoConvert     =   -1
                     ErrorBeep       =   0
                     MaxLength       =   4
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
                  Begin TDBText6Ctl.TDBText tdbtDual 
                     Height          =   315
                     Left            =   3360
                     TabIndex        =   13
                     Top             =   270
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":69FE
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":6A6A
                     Key             =   "frmManPlanCuentas.frx":6A88
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   8388608
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
                     Format          =   "9"
                     FormatMode      =   0
                     AutoConvert     =   -1
                     ErrorBeep       =   0
                     MaxLength       =   4
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
                  Begin TDBText6Ctl.TDBText tdbtFlujoEfectivo 
                     Height          =   315
                     Left            =   2595
                     TabIndex        =   16
                     Top             =   1380
                     Visible         =   0   'False
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     Caption         =   "frmManPlanCuentas.frx":6ADA
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmManPlanCuentas.frx":6B46
                     Key             =   "frmManPlanCuentas.frx":6B64
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   8388608
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
                     Format          =   "9"
                     FormatMode      =   0
                     AutoConvert     =   -1
                     ErrorBeep       =   0
                     MaxLength       =   4
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
                  Begin VB.Label Label15 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Flujo de Efectivo "
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   120
                     TabIndex        =   172
                     Top             =   1380
                     Visible         =   0   'False
                     Width           =   2385
                  End
                  Begin VB.Label Label9 
                     Height          =   255
                     Left            =   4125
                     TabIndex        =   55
                     Top             =   300
                     Width           =   195
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Est. de Result. Integ.- Natu."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   105
                     TabIndex        =   53
                     Top             =   1050
                     Width           =   2385
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Est. de Result. Integ.- func."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   105
                     TabIndex        =   52
                     Top             =   690
                     Width           =   2385
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Estado de Situacion Finan."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   105
                     TabIndex        =   51
                     Top             =   315
                     Width           =   2385
                  End
               End
               Begin TrueOleDBList70.TDBCombo tdbcOperaTC 
                  Height          =   300
                  Left            =   10260
                  TabIndex        =   11
                  Tag             =   "_"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   2310
                  _ExtentX        =   4075
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=688"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=609"
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
                  _PropDict       =   $"frmManPlanCuentas.frx":6BB6
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
               Begin TDBText6Ctl.TDBText TxtOrdCentrComp 
                  Height          =   315
                  Left            =   2745
                  TabIndex        =   157
                  Top             =   4560
                  Visible         =   0   'False
                  Width           =   750
                  _Version        =   65536
                  _ExtentX        =   1323
                  _ExtentY        =   556
                  Caption         =   "frmManPlanCuentas.frx":6C3D
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmManPlanCuentas.frx":6CA9
                  Key             =   "frmManPlanCuentas.frx":6CC7
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   8388608
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
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  MultiLine       =   0
                  ScrollBars      =   0
                  PasswordChar    =   ""
                  AllowSpace      =   -1
                  Format          =   "9"
                  FormatMode      =   0
                  AutoConvert     =   -1
                  ErrorBeep       =   0
                  MaxLength       =   4
                  LengthAsByte    =   0
                  Text            =   "0"
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
               Begin TDBText6Ctl.TDBText TxtOrdVta 
                  Height          =   315
                  Left            =   7575
                  TabIndex        =   159
                  Top             =   4560
                  Visible         =   0   'False
                  Width           =   750
                  _Version        =   65536
                  _ExtentX        =   1323
                  _ExtentY        =   556
                  Caption         =   "frmManPlanCuentas.frx":6D19
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmManPlanCuentas.frx":6D85
                  Key             =   "frmManPlanCuentas.frx":6DA3
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   8388608
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
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  MultiLine       =   0
                  ScrollBars      =   0
                  PasswordChar    =   ""
                  AllowSpace      =   -1
                  Format          =   "9"
                  FormatMode      =   0
                  AutoConvert     =   -1
                  ErrorBeep       =   0
                  MaxLength       =   4
                  LengthAsByte    =   0
                  Text            =   "0"
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
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Orden de Centr. de Comp. :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   150
                  TabIndex        =   160
                  Top             =   4590
                  Visible         =   0   'False
                  Width           =   2520
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Orden de Centr. de Vta. :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4980
                  TabIndex        =   158
                  Top             =   4635
                  Visible         =   0   'False
                  Width           =   2520
               End
            End
            Begin VB.Frame fraTexto 
               Height          =   4035
               Left            =   -74820
               TabIndex        =   125
               Top             =   450
               Width           =   8400
               Begin VB.Label lblMensaje 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   $"frmManPlanCuentas.frx":6DF5
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1725
                  Left            =   1170
                  TabIndex        =   126
                  Top             =   1125
                  Width           =   6315
               End
            End
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   1830
            TabIndex        =   4
            Tag             =   "_"
            Top             =   375
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "frmManPlanCuentas.frx":6E9E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlanCuentas.frx":6F0A
            Key             =   "frmManPlanCuentas.frx":6F28
            BackColor       =   16514043
            EditMode        =   0
            ForeColor       =   8388608
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
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   1830
            TabIndex        =   6
            Tag             =   "_"
            Top             =   735
            Width           =   5655
            _Version        =   65536
            _ExtentX        =   9975
            _ExtentY        =   556
            Caption         =   "frmManPlanCuentas.frx":6F7A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlanCuentas.frx":6FE6
            Key             =   "frmManPlanCuentas.frx":7004
            BackColor       =   16514043
            EditMode        =   0
            ForeColor       =   8388608
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
         Begin VB.Label lblMante 
            AutoSize        =   -1  'True
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
            Height          =   225
            Left            =   360
            TabIndex        =   31
            Top             =   105
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
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
            Left            =   360
            TabIndex        =   27
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
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
            Left            =   360
            TabIndex        =   26
            Top             =   420
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6075
         Left            =   -74865
         TabIndex        =   24
         Top             =   450
         Width           =   10095
         Begin TrueOleDBGrid70.TDBGrid tdbgCuentas 
            Height          =   4320
            Left            =   105
            TabIndex        =   3
            Top             =   1605
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   7620
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Cuenta"
            Columns(0).DataField=   "Pla_cCuentaContable"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre de Cuenta"
            Columns(1).DataField=   "Pla_cNombreCuenta"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   16
            Columns(2)._MaxComboItems=   5
            Columns(2).ValueItems(0)._DefaultItem=   0
            Columns(2).ValueItems(0).Value=   "S"
            Columns(2).ValueItems(0).Value.vt=   8
            Columns(2).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(2).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(2).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(2).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(2).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(2).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(2).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(2).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(2).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(2).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(2).ValueItems(0).DisplayValue.vt=   9
            Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems(1)._DefaultItem=   0
            Columns(2).ValueItems(1).Value=   "N"
            Columns(2).ValueItems(1).Value.vt=   8
            Columns(2).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(2).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(2).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(2).ValueItems(1).DisplayValue.vt=   9
            Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems.Count=   2
            Columns(2).Caption=   "Titulo"
            Columns(2).DataField=   "Pla_cTitulo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tip Cta"
            Columns(3).DataField=   "Pla_cTipoCta"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Balance"
            Columns(4).DataField=   "Pla_cCptoBG"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Función"
            Columns(5).DataField=   "Pla_cCptoResFun"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Naturaleza"
            Columns(6).DataField=   "Pla_cCptoResNat"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Tipo"
            Columns(7).DataField=   "Ten_cTipoEntidad"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Presup"
            Columns(8).DataField=   "DescriPresup"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   16
            Columns(9)._MaxComboItems=   5
            Columns(9).ValueItems(0)._DefaultItem=   0
            Columns(9).ValueItems(0).Value=   "1"
            Columns(9).ValueItems(0).Value.vt=   8
            Columns(9).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(9).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(9).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(9).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(9).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(9).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(9).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(9).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(9).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(9).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(9).ValueItems(0).DisplayValue.vt=   9
            Columns(9).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(9).ValueItems(1)._DefaultItem=   0
            Columns(9).ValueItems(1).Value=   "0"
            Columns(9).ValueItems(1).Value.vt=   8
            Columns(9).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(9).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(9).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(9).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(9).ValueItems(1).DisplayValue.vt=   9
            Columns(9).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(9).ValueItems.Count=   2
            Columns(9).Caption=   "CCosto"
            Columns(9).DataField=   "Pla_cCentroCosto"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   16
            Columns(10)._MaxComboItems=   5
            Columns(10).ValueItems(0)._DefaultItem=   0
            Columns(10).ValueItems(0).Value=   "1"
            Columns(10).ValueItems(0).Value.vt=   8
            Columns(10).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(10).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(10).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(10).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(10).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(10).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(10).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(10).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(10).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(10).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(10).ValueItems(0).DisplayValue.vt=   9
            Columns(10).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(10).ValueItems(1)._DefaultItem=   0
            Columns(10).ValueItems(1).Value=   "0"
            Columns(10).ValueItems(1).Value.vt=   8
            Columns(10).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(10).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(10).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(10).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(10).ValueItems(1).DisplayValue.vt=   9
            Columns(10).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(10).ValueItems.Count=   2
            Columns(10).Caption=   "Prov"
            Columns(10).DataField=   "Pla_cProvision"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   16
            Columns(11)._MaxComboItems=   5
            Columns(11).ValueItems(0)._DefaultItem=   0
            Columns(11).ValueItems(0).Value=   "1"
            Columns(11).ValueItems(0).Value.vt=   8
            Columns(11).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(11).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(11).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(11).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(11).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(11).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(11).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(11).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(11).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(11).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(11).ValueItems(0).DisplayValue.vt=   9
            Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(11).ValueItems(1)._DefaultItem=   0
            Columns(11).ValueItems(1).Value=   "0"
            Columns(11).ValueItems(1).Value.vt=   8
            Columns(11).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(11).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(11).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(11).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(11).ValueItems(1).DisplayValue.vt=   9
            Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(11).ValueItems.Count=   2
            Columns(11).Caption=   "Doc"
            Columns(11).DataField=   "Pla_cDocumento"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   16
            Columns(12)._MaxComboItems=   5
            Columns(12).ValueItems(0)._DefaultItem=   0
            Columns(12).ValueItems(0).Value=   "1"
            Columns(12).ValueItems(0).Value.vt=   8
            Columns(12).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(12).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(12).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(12).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(12).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(12).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(12).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(12).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(12).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(12).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(12).ValueItems(0).DisplayValue.vt=   9
            Columns(12).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(12).ValueItems(1)._DefaultItem=   0
            Columns(12).ValueItems(1).Value=   "0"
            Columns(12).ValueItems(1).Value.vt=   8
            Columns(12).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(12).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(12).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(12).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(12).ValueItems(1).DisplayValue.vt=   9
            Columns(12).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(12).ValueItems.Count=   2
            Columns(12).Caption=   "Detrac"
            Columns(12).DataField=   "pla_cdetraccion"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   16
            Columns(13)._MaxComboItems=   5
            Columns(13).ValueItems(0)._DefaultItem=   0
            Columns(13).ValueItems(0).Value=   "1"
            Columns(13).ValueItems(0).Value.vt=   8
            Columns(13).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(13).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(13).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(13).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(13).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(13).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(13).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(13).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(13).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(13).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(13).ValueItems(0).DisplayValue.vt=   9
            Columns(13).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(13).ValueItems(1)._DefaultItem=   0
            Columns(13).ValueItems(1).Value=   "0"
            Columns(13).ValueItems(1).Value.vt=   8
            Columns(13).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(13).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(13).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(13).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(13).ValueItems(1).DisplayValue.vt=   9
            Columns(13).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(13).ValueItems.Count=   2
            Columns(13).Caption=   "Reten"
            Columns(13).DataField=   "pla_cretencion"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   16
            Columns(14)._MaxComboItems=   5
            Columns(14).ValueItems(0)._DefaultItem=   0
            Columns(14).ValueItems(0).Value=   "1"
            Columns(14).ValueItems(0).Value.vt=   8
            Columns(14).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(14).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(14).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(14).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(14).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(14).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(14).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(14).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(14).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(14).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(14).ValueItems(0).DisplayValue.vt=   9
            Columns(14).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(14).ValueItems(1)._DefaultItem=   0
            Columns(14).ValueItems(1).Value=   "0"
            Columns(14).ValueItems(1).Value.vt=   8
            Columns(14).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(14).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(14).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(14).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(14).ValueItems(1).DisplayValue.vt=   9
            Columns(14).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(14).ValueItems.Count=   2
            Columns(14).Caption=   "Percep"
            Columns(14).DataField=   "pla_cpercepcion"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "FlujoEfec"
            Columns(15).DataField=   "Pla_cCptoEFE"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   16
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=16"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7514"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7435"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=873"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=794"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=529"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=529"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1191"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1111"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=529"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1191"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1111"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=529"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1508"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1429"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=529"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=820"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=741"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=529"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1058"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=979"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=532"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1323"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1244"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=529"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(61)=   "Column(10).Width=1270"
            Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1191"
            Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=529"
            Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(67)=   "Column(11).Width=873"
            Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=794"
            Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=529"
            Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(73)=   "Column(12).Width=1138"
            Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=1058"
            Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(77)=   "Column(12)._ColStyle=529"
            Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(79)=   "Column(13).Width=1164"
            Splits(0)._ColumnProps(80)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(81)=   "Column(13)._WidthInPix=1085"
            Splits(0)._ColumnProps(82)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(83)=   "Column(13)._ColStyle=529"
            Splits(0)._ColumnProps(84)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(85)=   "Column(14).Width=1032"
            Splits(0)._ColumnProps(86)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(87)=   "Column(14)._WidthInPix=953"
            Splits(0)._ColumnProps(88)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(89)=   "Column(14)._ColStyle=529"
            Splits(0)._ColumnProps(90)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(91)=   "Column(15).Width=1402"
            Splits(0)._ColumnProps(92)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(93)=   "Column(15)._WidthInPix=1323"
            Splits(0)._ColumnProps(94)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(95)=   "Column(15)._ColStyle=529"
            Splits(0)._ColumnProps(96)=   "Column(15).Order=16"
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
            MultiSelect     =   2
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
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=94,.parent=13,.alignment=2"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=91,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=92,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=93,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=87,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=88,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=89,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=86,.parent=13,.alignment=2"
            _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=83,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=84,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=85,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=98,.parent=13,.alignment=2"
            _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
            _StyleDefs(102) =   "Named:id=33:Normal"
            _StyleDefs(103) =   ":id=33,.parent=0"
            _StyleDefs(104) =   "Named:id=34:Heading"
            _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(106) =   ":id=34,.wraptext=-1"
            _StyleDefs(107) =   "Named:id=35:Footing"
            _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(109) =   "Named:id=36:Selected"
            _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(111) =   "Named:id=37:Caption"
            _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(113) =   "Named:id=38:HighlightRow"
            _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(115) =   "Named:id=39:EvenRow"
            _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(117) =   "Named:id=40:OddRow"
            _StyleDefs(118) =   ":id=40,.parent=33"
            _StyleDefs(119) =   "Named:id=41:RecordSelector"
            _StyleDefs(120) =   ":id=41,.parent=34"
            _StyleDefs(121) =   "Named:id=42:FilterBar"
            _StyleDefs(122) =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtCodigoBus 
            Height          =   315
            Left            =   1500
            TabIndex        =   0
            Top             =   630
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "frmManPlanCuentas.frx":7056
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlanCuentas.frx":70C2
            Key             =   "frmManPlanCuentas.frx":70E0
            BackColor       =   16514043
            EditMode        =   0
            ForeColor       =   8388608
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1500
            TabIndex        =   2
            Top             =   990
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   556
            Caption         =   "frmManPlanCuentas.frx":7132
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlanCuentas.frx":719E
            Key             =   "frmManPlanCuentas.frx":71BC
            BackColor       =   16514043
            EditMode        =   0
            ForeColor       =   8388608
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
         Begin TDBText6Ctl.TDBText tdbtTituloBus 
            Height          =   315
            Left            =   5340
            TabIndex        =   1
            Top             =   630
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            Caption         =   "frmManPlanCuentas.frx":720E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManPlanCuentas.frx":727A
            Key             =   "frmManPlanCuentas.frx":7298
            BackColor       =   16514043
            EditMode        =   0
            ForeColor       =   8388608
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   1
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
         Begin VB.Line Line3 
            BorderColor     =   &H80000003&
            X1              =   180
            X2              =   9840
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   195
            X2              =   9855
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Label lblAnio 
            AutoSize        =   -1  'True
            Caption         =   "AÑO:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   8715
            TabIndex        =   38
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cuenta :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   360
            TabIndex        =   37
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Nombre :"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   36
            Top             =   990
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filtrar Datos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   35
            Top             =   210
            Width           =   1200
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Titulo :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   4080
            TabIndex        =   30
            Top             =   630
            Width           =   1020
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4230
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":72EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":76C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":7A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":7E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":8252
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":862C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":8A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":8DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManPlanCuentas.frx":9DFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   264
      Left            =   12
      TabIndex        =   153
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
Attribute VB_Name = "frmManPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------
' Creado por    :  Miguel Angel Lopez Sanabria
' Descripción   :  Realiza mantenimiento a la tabla plan de Cuentas
'                  y define sus parámetros.
' -----------------------------------------------------------------------------
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lArrDestino() As Variant    ' *** Arreglo para los destinos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lFlag As String
Dim Control As String           ' *** Para busqueda
Dim lrsPlanCta As ADODB.Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim auxDestino As Boolean
Dim sw As Boolean
Dim gsGrupo As String
Dim nFilas As Integer
Dim gsDigitosCtaDetalle As Integer
Dim RegPI As String

Public Property Let Grupo(ByVal Grupo As String)
 gsGrupo = Grupo
End Property

Private Sub chkCentroCosto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then pSendKeys "{TAB}"
 If KeyAscii = 48 Then chkCentroCosto.Value = 0
 If KeyAscii = 49 Then chkCentroCosto.Value = 1
End Sub

'Private Sub ChkConsPDT601_Click()
'    If CmbRPI.Text <> "" And ChkConsPDT601.Value = 1 Then
'        If MsgBox("Si desea considerar la presente cuenta para efectos de PDT 601-PLAME, no seleccione la opción de Regimen Pensionario de Independientes ", vbInformation + vbOKOnly, gsNombreModulo) = vbOK Then
'            ChkConsPDT601.Value = 0
'        End If
'    End If
'End Sub

Private Sub chkDocumento_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then pSendKeys "{TAB}"
 If KeyAscii = 48 Then chkDocumento.Value = 0
 If KeyAscii = 49 Then chkDocumento.Value = 1
End Sub
Private Sub chkFun_Click()
    Call HabilitarCamposFuncion
End Sub

Private Sub chkNat_Click()
    Call HabilitarCamposNaturaleza
End Sub

Private Sub chkProvision_Click()
    'chkDocumento.Enabled = True
    If chkProvision.Value = vbChecked Then
        chkDocumento.Value = vbChecked
        chkDocumento.Enabled = False
    Else
        chkDocumento.Enabled = True
    End If
End Sub

Private Sub chkProvision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
    If KeyAscii = 48 Then chkProvision.Value = 0
    If KeyAscii = 49 Then chkProvision.Value = 1
End Sub

Private Sub chkTitulo_Click()
    On Error GoTo serror:
    If chkTitulo.Value = vbChecked Then
        fraTitulos.Visible = True
        fraNoTitulos.Visible = False
        
        Call ActivarNoTitulos(vbUnchecked)
        Call ActivarNoTituloCta7(vbUnchecked)
        Call ActivarNoTituloCta456(vbUnchecked)
        Call ActivarNoTituloCta8(vbUnchecked)
        
    Else
        fraTitulos.Visible = False
        fraNoTitulos.Visible = True
        
        Call ActivarTitulos(vbUnchecked)
        Call ActivarNoTituloCta7(vbUnchecked)
        Call ActivarNoTituloCta456(vbUnchecked)
        Call ActivarNoTituloCta8(vbUnchecked)
        
    End If
    
    If Left(tdbtCodigo.Text, 3) = "711" Or Left(tdbtCodigo.Text, 3) = "715" Or Left(tdbtCodigo.Text, 1) = "9" Then
        Me.fraTexto.Visible = False
        fraTitulos.Visible = True
    End If
    
   Exit Sub
serror:
    Mensajes Err.Description
End Sub

Private Sub chkTitulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
    If KeyAscii = 48 Then chkTitulo.Value = 0
    If KeyAscii = 49 Then chkTitulo.Value = 1
End Sub

'Private Sub CmbRPI_Click()
'    If ChkConsPDT601.Value = 1 And CmbRPI.Text <> "" Then
'        If MsgBox("Si desea considerar la presente cuenta para efectos de Retención de Independientes, no seleccione la opción Considerar PDT 601-PLAME", vbInformation + vbOKOnly, gsNombreModulo) = vbOK Then
'            CmbRPI.ListIndex = 0
'            chkDocumento.Value = 0
'        End If
'    ElseIf CmbRPI.Text <> "" Then
'        chkDocumento.Value = 1
'    ElseIf CmbRPI.Text = "" Then
'        chkDocumento.Value = 0
'    End If
'End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
End Sub
Private Function ValidaCuentas() As Boolean
    ValidaCuentas = False
    If tdbtCodGanancia = "" Then Mensajes "Ingrese Cuenta de Ganancia por Diferencia de Cambio ", vbInformation: Exit Function

    If tdbtCodPerdida = "" Then Mensajes "Ingrese Cuenta de Perdida por Diferencia de Cambio ", vbInformation: Exit Function
    
    If tdbtCodRedondeoG = "" Then Mensajes "Ingrese Cuenta de Ganancia por Redondeo ", vbInformation: Exit Function
    
    If tdbtCodRedondeoP = "" Then Mensajes "Ingrese Cuenta de Perdida por Redondeo", vbInformation: Exit Function
    
    'Valida que no sean iguales las cuentas de G y P por Diferencia de cambio
    If Trim(tdbtCodGanancia) <> "" And Trim(tdbtCodGanancia) = Trim(tdbtCodPerdida) Then
       Mensajes "La cuentas por Diferencia de Cambio no deben ser iguales. Verifique...", vbInformation
       Exit Function
    End If
    
    'Valida que no sean iguales las cuentas de G y P por Redondero
    If Trim(tdbtCodRedondeoG) <> "" And Trim(tdbtCodRedondeoG) = Trim(tdbtCodRedondeoP) Then
       Mensajes "La cuentas por Redondeo no deben ser iguales. Verifique...", vbInformation
       Exit Function
    End If
    '------------------------------------------------
    If InStr(1, tdbtCodRedondeoP.Text & " " & tdbtCodGanancia.Text & " " & tdbtCodPerdida.Text, tdbtCodRedondeoG.Text) <> 0 Then
       Mensajes "La cuenta " & tdbtCodRedondeoG.Text & " no debe repetirse. Verifique...", vbInformation
       pSetFocus tdbtCodRedondeoP
       Exit Function
    End If
    
    If InStr(1, tdbtCodRedondeoG.Text & " " & tdbtCodGanancia.Text & " " & tdbtCodPerdida.Text, tdbtCodRedondeoP.Text) <> 0 Then
       Mensajes "La cuenta " & tdbtCodRedondeoP.Text & " no debe repetirse. Verifique...", vbInformation
       pSetFocus tdbtCodRedondeoG
       Exit Function
    End If
    
    If InStr(1, tdbtCodRedondeoG.Text & " " & tdbtCodRedondeoP.Text & " " & tdbtCodPerdida.Text, tdbtCodGanancia.Text) <> 0 Then
       Mensajes "La cuenta " & tdbtCodGanancia.Text & " no debe repetirse. Verifique...", vbInformation
       pSetFocus tdbtCodRedondeoG
       Exit Function
    End If
    
    If InStr(1, tdbtCodRedondeoG.Text & " " & tdbtCodRedondeoP.Text & " " & tdbtCodGanancia.Text, tdbtCodPerdida.Text) <> 0 Then
       Mensajes "La cuenta " & tdbtCodPerdida.Text & " no debe repetirse. Verifique...", vbInformation
       pSetFocus tdbtCodRedondeoG
       Exit Function
    End If
    
    ValidaCuentas = True
End Function

Private Sub cmdActualizar_Click()
    ' *** Actualizando las cuentas del Plan Contable
    Dim clsMante As clsMantoTablas
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    
    If ValidaCuentas = False Then Exit Sub
    
    Set clsMante = New clsMantoTablas
    
    ReDim lArrDestino(9) As Variant
    lArrDestino(0) = "DIFCAMBIO"        ' Accion
    lArrDestino(1) = gsEmpresa          ' Empresa
    lArrDestino(2) = gsAnio             ' Anio
    lArrDestino(3) = ""                 ' Mes
    lArrDestino(4) = ""                 ' Cuenta
    lArrDestino(5) = ""                 ' Secuencia
    lArrDestino(6) = tdbtCodGanancia    ' DestinoDebe
    lArrDestino(7) = tdbtCodPerdida     ' DestinoHaber
    lArrDestino(8) = 0                  ' Porcentaje
    lArrDestino(9) = gsUsuario          ' Usuario
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrDestino(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    ReDim lArrDestino(9) As Variant
    lArrDestino(0) = "REDONDEO"        ' Accion
    lArrDestino(1) = gsEmpresa          ' Empresa
    lArrDestino(2) = gsAnio             ' Anio
    lArrDestino(3) = ""                 ' Mes
    lArrDestino(4) = ""                 ' Cuenta
    lArrDestino(5) = ""                 ' Secuencia
    lArrDestino(6) = tdbtCodRedondeoG    ' DestinoDebe
    lArrDestino(7) = tdbtCodRedondeoP   ' DestinoHaber
    lArrDestino(8) = 0                  ' Porcentaje
    lArrDestino(9) = gsUsuario          ' Usuario
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrDestino(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    ' *** Mensajes y Desactivar controles
    Mensajes "Los datos se actualizaron con exito", vbInformation
    tdbtCodGanancia.ReadOnly = True
    tdbtCodPerdida.ReadOnly = True
    tdbtCodRedondeoG.ReadOnly = True
    tdbtCodRedondeoP.ReadOnly = True
 
    cmdActualizar.Enabled = False
    cmdEditar.Enabled = True
    TabMantenimiento False
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub cmdEditar_Click()
    TabMantenimiento True, True
    sstPerfiles.TabEnabled(0) = False
    sstPerfiles.TabEnabled(1) = False
    sstPerfiles.Tab = 2
    
    tdbtCodGanancia.ReadOnly = False
    tdbtCodPerdida.ReadOnly = False
    tdbtCodRedondeoG.ReadOnly = False
    tdbtCodRedondeoP.ReadOnly = False
    pSetFocus tdbtCodGanancia
    cmdActualizar.Enabled = True
    cmdEditar.Enabled = False
End Sub

Private Sub cmdEliminarDestino_Click()
    Dim Fila As Integer
    If tdblDestinoAux.ListCount > 0 Then
        Fila = tdblDestinoAux.Row
        If tdblDestinoAux.Row <> -1 Then
            tdblDestinoAux.RemoveItem tdblDestinoAux.Bookmark
            ''Call EliminarDatosLista
        Else
            Mensajes "Seleccione el item a eliminar", vbInformation
        End If
    End If
End Sub
'Private Sub EliminarDatosLista()
'    Dim i As Integer
'    ' *** Insertar Registro a la Lista de Destino
'    With tdblDestinoAux
'        If lTipoMnt = "INSERTAR" Then
'            ' *** Si es nuevo
'            For i = 0 To .ListCount - 1
'                .Bookmark = i
'                If Trim(.Columns(1).Value) = Trim(tdblDestinoAux.Columns(0).Value) And _
'                    Trim(.Columns(2).Value) = Trim(tdblDestinoAux.Columns(1).Value) And .Columns(4).Value = tdblDestinoAux.Columns(3).Value Then
'                    tdblDestinoAux.RemoveItem i
'                    i = i - 1
'                End If
'            Next
'        Else
'            For i = 0 To .ListCount
'                .Bookmark = i
'                If Trim(.Columns(0).Value) = tdbcMes.BoundText And Trim(.Columns(1).Value) = Trim(tdblDestinoAux.Columns(0).Value) And _
'                    Trim(.Columns(2).Value) = Trim(tdblDestinoAux.Columns(1).Value) And .Columns(4).Value = tdblDestinoAux.Columns(3).Value Then
'                    tdblDestinoAux.RemoveItem i
'                    Exit For
'                End If
'            Next
'        End If
'    End With
'    tdblDestinoAux.ReBind
'End Sub
Private Function EsCtaDestino() As Boolean
    Dim sqlCta As String
    
    EsCtaDestino = False
    
    If Left(CE(tdbtCodigo), 1) = "9" Or Left(CE(tdbtCodigo), 1) = "6" Then
        EsCtaDestino = True
    End If
    
    
    'sqlCta = "SELECT count(*) from CNA_CTAS_CONDESTINO WHERE Emp_cCodigo = '" & gsEmpresa & "' "
    'sqlCta = sqlCta + "AND CdE_cClase = '" & Left(CE(tdbtCodigo), 1) & "' AND Cde_cEstado = 'A' "
    'If ExisteDato(sqlCta) Then EsCtaDestino = True
    ' ***
End Function
Private Sub cmdInsertar_Click()
    Dim i As Integer
    ' *** Verificar q sea tipo de cuenta con destino
    If CE(tdbtCtaDestino.Text) = "" Then
        Mensajes "Ingrese una cuenta contable de destino. Verificar...", vbInformation
        pSetFocus tdbtCtaDestino
        Exit Sub
    End If
    
    If chkTitulo = 1 Then
        Mensajes "Cuenta es de titulo. Verificar...", vbInformation
        Exit Sub
    End If
   
    If Not EsCtaDestino Then
        Mensajes "Esta cuenta no esta seteada para tener destino", vbInformation
        Exit Sub
    End If
    
    ' *** Q el codigo sea diferente de nada
    If TextoLleno(tdbtCtaDestino, "Cuenta") = False Then Exit Sub
    If tdbnPorc = 0 Then
        Mensajes "Ingrese una cantidad en porcentaje diferente a 0", vbInformation
        Exit Sub
    End If
        
    If Left(tdbtCodigo.Text, 1) = Left(tdbtCtaDestino.Text, 1) And EsCtaDestino And Left(cmbTipo.Text, 1) = "D" Then
        Mensajes "La cuenta no debe ser de la misma clase " & Salto(1) & "de la cuenta de origen " & Left(tdbtCodigo.Text, 1) & " en el DEBE"
        tdbtCtaDestino.Text = ""
        tdbtNombreDestino.Text = ""
        pSetFocus tdbtCtaDestino
        Exit Sub
    End If
                
    ' *** Insertar Registro a la Lista de Destino
    If tdblDestinoAux.ListCount < 8 Then auxDestino = True
    If Me.cmbTipo = "Debe" Then
        tdblDestinoAux.AddItem tdbtCtaDestino & "; ; " & tdbtNombreDestino & " ;" & tdbnPorc
'        If lTipoMnt = "INSERTAR" Then
'            ' *** Si es nuevo
'            For i = 0 To 12
'                tdblDestino.AddItem Format(i, "00") & "; " & tdbtCtaDestino & "; ; " & tdbtNombreDestino & " ;" & tdbnPorc
'            Next
'        Else
'            If auxDestino = True Then
'                For i = 0 To 12
'                    tdblDestino.AddItem Format(i, "00") & "; " & tdbtCtaDestino & "; ; " & tdbtNombreDestino & " ;" & tdbnPorc
'                Next
'            Else
'                tdblDestino.AddItem Me.tdbcMes.BoundText & "; " & tdbtCtaDestino & "; ; " & tdbtNombreDestino & " ;" & tdbnPorc
'            End If
'        End If
    Else
        tdblDestinoAux.AddItem " ; " & tdbtCtaDestino & "; " & tdbtNombreDestino & " ;" & tdbnPorc
'        If lTipoMnt = "INSERTAR" Then
'            ' *** Si es nuevo
'            For i = 0 To 12
'                tdblDestino.AddItem Format(i, "00") & " ; ; " & tdbtCtaDestino & "; " & tdbtNombreDestino & " ;" & tdbnPorc
'            Next
'        Else
'            If auxDestino = True Then
'                For i = 0 To 12
'                    tdblDestino.AddItem Format(i, "00") & " ; ; " & tdbtCtaDestino & "; " & tdbtNombreDestino & " ;" & tdbnPorc
'                Next
'            Else
'                tdblDestino.AddItem tdbcMes.BoundText & " ; ; " & tdbtCtaDestino & "; " & tdbtNombreDestino & " ;" & tdbnPorc
'            End If
'        End If
    End If
    ' *** Limpiar lo q se escribio previamente
    pSetFocus tdbtCtaDestino
    If cmbTipo.Text = "Debe" Then
        cmbTipo.Text = "Haber"
    Else
        cmbTipo.Text = "Debe"
    End If
    
    Me.tdbtCtaDestino = ""
    Me.tdbtNombreDestino = ""
    tdbnPorc = 100
    ' ***
End Sub
Private Function VerificarPorcDestino() As Boolean
    Dim debePor As Double
    Dim haberPor As Double
    Dim i As Integer
    
    debePor = 0: haberPor = 0
    VerificarPorcDestino = True
    For i = 0 To Me.tdblDestinoAux.ListCount - 1
        tdblDestinoAux.Bookmark = i
        If Trim(tdblDestinoAux.Columns(0).Value) <> "" Then
            debePor = debePor + tdblDestinoAux.Columns(3).Value
        Else
            haberPor = haberPor + tdblDestinoAux.Columns(3).Value
        End If
    Next
    If debePor <> 100 And debePor <> 0 Then
        tdblDestinoAux.Refresh
        
        Mensajes "Destino del debe diferente de 100. Verificar...", vbInformation
                tdblDestinoAux.Refresh
        VerificarPorcDestino = False
    End If
    If haberPor <> 100 And haberPor <> 0 Then
            tdblDestinoAux.Refresh
        Mensajes "Destino del haber diferente de 100. Verificar...", vbInformation
                tdblDestinoAux.Refresh
        VerificarPorcDestino = False
    End If
End Function
Private Sub Form_Activate()
    ' *** Activar el año de trabajo
    If sw = True Then Exit Sub
    lblanio.Caption = "AÑO: " & gsAnio
    
    OptAct.Enabled = True
    OptPas.Enabled = True
    chkNat.Enabled = False
    chkFun.Enabled = False
    OptAct.Value = True
    
    tdbcPatrimomio.Enabled = False
    tdbtResFunc.Enabled = False
    tdbtResNatu.Enabled = False
    sw = True

    If gsDigitosCtaDetalle < 4 Then
        Mensajes "Configure el número de digitos de la cuenta de detalle" & Salto(1) & "se activara el formulario de configuracion inicial"
        Unload Me
        
        frmManSubDiarioTDoc.Grupo = BuscaArray("mnuParamIniciales")
        frmManSubDiarioTDoc.Show
        pSetFocus frmManSubDiarioTDoc.tdbnNdigitos
    
        Exit Sub
    End If
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    
    Select Case KeyCode
        Case 27:
            If sstPerfiles.TabEnabled(1) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo: Exit Sub
        Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos: Exit Sub
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar: Exit Sub
        Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar: Exit Sub
        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar: Exit Sub
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir: Exit Sub
        
    End Select
    ' ***
End Sub

Private Sub HabilitarCampos()
    Call HabilitarCamposFuncion
    Call HabilitarCamposNaturaleza
    Call HabilitarCamposActivoPasivo
End Sub

Private Sub HabilitarCamposFuncion()
    If chkFun.Value = vbChecked Then
        ActivarControl tdbtResFunc, True
    Else
        ActivarControl tdbtResFunc, False
        tdbtResFunc.Text = ""
    End If
End Sub

Private Sub HabilitarCamposNaturaleza()
    If chkNat.Value = vbChecked Then
        ActivarControl tdbtResNatu, True
    Else
        ActivarControl tdbtResNatu, False
        tdbtResNatu.Text = ""
    End If
End Sub

Private Sub HabilitarCamposActivoPasivo()
    If OptAct.Value = True Or OptPas.Value = True Then
        ActivarControl tdbtBalance, True
        ActivarControl tdbtDual, True
    Else
        ActivarControl tdbtBalance, False
        ActivarControl tdbtDual, False
        tdbtBalance.Text = ""
        tdbtDual.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Me.Hide
    DoEvents
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    gsDigitosCtaDetalle = NE(BuscaValorEnOp("054"))
    
    Call Centrar_form(Me)
    Call CargaTabla
    Call LlenaCombos
    
    ' *** Seteando algunas variables
    lRegElim = False
    lTipoMnt = "INSERTAR"
    lFlag = "1"
    sstPerfiles.TabEnabled(1) = False
    sstPerfiles.Tab = 0
    
    EstadoLDOri = "1"
    
    cmbTipo.ListIndex = 0
    
    ' *** Traer los datos de las Cuentas por Diferencia de Cambio
    tdbtCodGanancia = CuentaCfgAuto("SEL_GAN")
    tdbtCodPerdida = CuentaCfgAuto("SEL_PER")
    tdbtCodRedondeoG = CuentaCfgAuto("SEL_REDONDEOG")
    tdbtCodRedondeoP = CuentaCfgAuto("SEL_REDONDEOP")
    
    tdbtCodGanancia_LostFocus
    tdbtCodPerdida_LostFocus
    tdbtCodRedondeoG_LostFocus
    tdbtCodRedondeoP_LostFocus
    sw = False
    ' ***
    Call LlenaComboMesActivo(tdbcMes)
    
    SeteaBarraHerramientas tbrOpciones, gsGrupo
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdEditar.Enabled = False
        Me.cmdActualizar.Enabled = False
        
    Else
        Me.cmdEditar.Enabled = True
        Me.cmdActualizar.Enabled = True
    End If
    
    Me.cmdActualizar.Enabled = False
    
    Call AsignaCodigosOPChecks
    
    Call Habilita39(False)
    
    lblMensaje.Caption = Salto(1) & "Esta sección de configuración de CUENTAS CONTABLES, " & _
                                    "esta diseñada solo para las cuentas de tipo TITULO de 2 digitos o " & _
                                    "para las cuentas de tipo DETALLE de " & gsDigitosCtaDetalle & " digitos"
    

    DoEvents
    Me.Show
End Sub

Private Function ValidaMovimientoEnCuenta() As Boolean
    If CE(Me.tdbtCodigo.Text) = "" Then ValidaMovimientoEnCuenta = True: Exit Function

    ValidaMovimientoEnCuenta = False
    
    'Validar la cuenta si tiene mov. y se cambia a Titulo
    'If chkTitulo.Value = 1 Then
        If VerificaCuentasMvtos(Me.tdbtCodigo) = True Then
           Mensajes "Esta cuenta tiene movimientos." & Salto(1) & "Elimine los movimientos contables de esta cuenta" & Salto(1) & _
                    "si desea modificar los PARAMETROS..." & Salto(2) & "Sólo podra modificar la opcion CONFIGURACION", vbInformation
                    
            sstParamatros.Tab = 2
           Exit Function
        End If
    'End If
    
    'Valida si cuenta tiene mov. de provisión. No debe permitir modificar el check de provisión
    'If chkProvision.Value = 0 Then
        If VerificaMovProvision(Me.tdbtCodigo) = True Then
           Mensajes "Esta cuenta tiene movimientos de provisión. Verifique los movimientos contables con esta cuenta...", vbInformation
           Exit Function
        End If
    'End If
    
    ValidaMovimientoEnCuenta = True
End Function

Private Sub Imprimir()
    frmRepPlanCuentas.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
       On Error GoTo serror
       sstPerfiles.Width = Me.Width - 200
       Frame1.Width = Me.Width - 500
       tdbgCuentas.Width = Me.Width - 800
       '-----------------------------------
       sstPerfiles.Height = Me.Height - 880 - 30
       Frame1.Height = Me.Height - 2000 + 400
       
       tdbgCuentas.Height = sstPerfiles.Height - tdbgCuentas.Top - 800
       tbrOpciones.Width = Me.Width
       '-----------------------------------
       Frame2.Width = sstPerfiles.Width - 300
       Frame2.Height = sstPerfiles.Height - 500
       
       sstPerfiles.Height = sstPerfiles.Height + 50
       sstParamatros.Width = sstPerfiles.Width - 500
       sstParamatros.Height = sstPerfiles.Height - sstParamatros.Top - 600
       
       Call FramesVisibles
       Call FramesVisiblesGral
       
       Call Centrar_Objeto(fraDifCambio, sstPerfiles, 0, 200)
       Call Centrar_Objeto(fradatos, sstParamatros, 0, 300)
       Call Centrar_Objeto(fraDestinoCuenta, sstParamatros)
       Call Centrar_Objeto(fraConfig, sstParamatros, 0, 200)
       Call Centrar_Objeto(fraTexto, sstParamatros, 0, 200)
              
    End If
    Exit Sub
serror:
    'Mensajes Err.Description
    
End Sub

Private Sub FramesVisiblesGral()
       Select Case sstPerfiles.Tab
        Case 0:
              fraDifCambio.Visible = False
        Case 1:
              fraDifCambio.Visible = False
        Case 2:
              fraDifCambio.Visible = True
       End Select
End Sub
Private Sub FramesVisibles()
       Select Case sstParamatros.Tab
        Case 0:
              fradatos.Visible = True
              fraDestinoCuenta.Visible = False
              fraConfig.Visible = False
              fraTexto.Visible = False
        Case 1:
              fradatos.Visible = False
              fraDestinoCuenta.Visible = True
              fraConfig.Visible = False
              fraTexto.Visible = False
        Case 2:
              fradatos.Visible = False
              fraDestinoCuenta.Visible = False
              fraConfig.Visible = True
              If Left(tdbtCodigo.Text, 3) = "711" Or Left(tdbtCodigo.Text, 3) = "715" Or Left(tdbtCodigo.Text, 1) = "9" Then
                 fraTexto.Visible = False
                 Me.fraTitulos.Visible = True
              Else
                Me.fraTexto.Visible = True
              End If
       End Select
End Sub

Private Sub OptAct_Click()
    Call HabilitarCamposActivoPasivo
End Sub

Private Sub OptPas_Click()
    Call HabilitarCamposActivoPasivo
End Sub

Private Sub sstParamatros_Click(PreviousTab As Integer)
    
    Call FramesVisibles
    
    If sstParamatros.Tab = 2 Then
        If Len(tdbtCodigo.Text) = 2 Or Len(tdbtCodigo.Text) = gsDigitosCtaDetalle Then
        
            If Len(tdbtCodigo.Text) = 2 And chkTitulo.Value = vbChecked Then
                fraConfig.Visible = True
            ElseIf Len(tdbtCodigo.Text) = gsDigitosCtaDetalle And chkTitulo.Value = vbUnchecked Then
                fraConfig.Visible = True
                Select Case Trim(Left(tdbtCodigo, 1))
                 Case "4", "6", "9"
                    Label13(10).Visible = True: chkCierreCtaCnfImp.Visible = True
                 Case Else
                    Label13(10).Visible = False: chkCierreCtaCnfImp.Visible = False
                End Select
            Else
                fraConfig.Visible = False
            End If
        Else
            fraConfig.Visible = False
        End If
'        If Trim(RegPI) <> "" Then CmbRPI.Text = RegPI
    
    End If
    
    If Left(tdbtCodigo.Text, 3) = "711" Or Left(tdbtCodigo.Text, 3) = "715" Or Left(tdbtCodigo.Text, 1) = "9" Then
      fraConfig.Visible = True
    End If
        
End Sub

Private Sub sstPerfiles_Click(PreviousTab As Integer)
    Call FramesVisiblesGral
    
    If sstPerfiles.Tab = 2 Then
        Call Centrar_Objeto(fraDifCambio, sstPerfiles, 0, 200)

        tbrOpciones.Buttons(1).Enabled = False
        tbrOpciones.Buttons(2).Enabled = False
        tbrOpciones.Buttons(3).Enabled = False
        tbrOpciones.Buttons(4).Enabled = False
        tbrOpciones.Buttons(5).Enabled = False
        tbrOpciones.Buttons(6).Enabled = False
        
    ElseIf sstPerfiles.Tab = 0 Then
        tbrOpciones.Buttons(1).Enabled = True
        tbrOpciones.Buttons(2).Enabled = True
        'tbrOpciones.Buttons(3).Enabled = True
        tbrOpciones.Buttons(4).Enabled = True
        tbrOpciones.Buttons(5).Enabled = True
        tbrOpciones.Buttons(6).Enabled = True
        
        SeteaBarraHerramientas tbrOpciones, gsGrupo
    End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: ManNuevo
        Case 2: VerDatos
        Case 3: Grabar
                If sstPerfiles.Tab = 0 Then tdbgCuentas.HighlightRowStyle = "HighlightRow"
                SeteaBarraHerramientas tbrOpciones, gsGrupo
        Case 4: Borrar
        Case 5: Editar
        Case 6: Imprimir
        Case 7: CancelarTB
                
    End Select
End Sub

Private Sub CancelarTB()
    Dim respuesta As String
    'If sstPerfiles.TabEnabled(1) = False Then ' *** Grabar
    If Me.tbrOpciones.Buttons(7).Image = 7 Then
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
        Unload Me
    Else
        respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
        If respuesta = vbYes Then
            Call Cancelar
            cmdActualizar.Enabled = False
            cmdEditar.Enabled = True
            
            tdbtCodGanancia = CuentaCfgAuto("SEL_GAN")
            tdbtCodPerdida = CuentaCfgAuto("SEL_PER")
            tdbtCodRedondeoG = CuentaCfgAuto("SEL_REDONDEOG")
            tdbtCodRedondeoP = CuentaCfgAuto("SEL_REDONDEOP")

            tdbtCodGanancia.ReadOnly = True
            tdbtCodPerdida.ReadOnly = True
            tdbtCodRedondeoG.ReadOnly = True
            tdbtCodRedondeoP.ReadOnly = True
            
            
            SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
            
            If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
                Me.cmdEditar.Enabled = False
            Else
                Me.cmdEditar.Enabled = True
            End If
            
        End If
       
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsPlanCta)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Set lrsPlanCta = New ADODB.Recordset
    Set lrsPlanCta.DataSource = Nothing
    
    '---------------------------
    ' BUSCA LAS CUENTA EN CONFIG OPERACIONES QUE NO SE ENCUENTRAN
    '  EN EL PLAN DE CUENTAS Y LAS ELIMINA
    ReDim lArrMnt(2) As Variant
    lArrMnt(0) = "LIMPIA_OPERACIONES"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    
    If clDatos.MantenimientoDeTablas(gsCadenaConexion, "spCn_ConsultaCuentas", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    '---------------------------
    
    sqlSp = "spCn_ConsultaCuentas 'SEL_ALL', '" & gsEmpresa & "', '" & gsAnio & "', ''"
    arrDatos = Array(sqlSp)
    
    
    Set lrsPlanCta = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    On Error Resume Next
    If Not lrsPlanCta Is Nothing Then
        If Not (lrsPlanCta.EOF And lrsPlanCta.BOF) Then
            lrsPlanCta.Sort = "Pla_cCuentaContable"
            tdbgCuentas.DataSource = lrsPlanCta
            nFilas = lrsPlanCta.RecordCount
        End If
    End If
    FiltrarRecordSet
End Sub

Private Function VerificaCuentasHijas(Cuenta As String) As Boolean

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    VerificaCuentasHijas = False
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant

    Dim sqlSp As String
    
    sqlSp = "spCn_ConsultaCuentas 'BUSCA_HIJOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & Cuenta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "Seleccione un registro", vbInformation
        Set rsArreglo = Nothing
        Exit Function
    End If
    If rsArreglo(0).Value > 1 Then VerificaCuentasHijas = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing

End Function

Private Function VerificaSiEsCuentasHija(Cuenta As String) As Boolean

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    VerificaSiEsCuentasHija = False
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant

    Dim sqlSp As String
    
    sqlSp = "spCn_ConsultaCuentas 'BUSCA_SI_ES_HIJO', '" & gsEmpresa & "', '" & gsAnio & "', '" & Cuenta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "Seleccione un registro", vbInformation
        Set rsArreglo = Nothing
        Exit Function
    End If
    If rsArreglo(0).Value > 1 Then VerificaSiEsCuentasHija = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing

End Function

Private Function VerificaSiEsCuentasTipo(Cuenta As String) As Boolean

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    VerificaSiEsCuentasTipo = False
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant

    Dim sqlSp As String
    
    sqlSp = "spCn_ConsultaCuentas 'BUSCA_SI_ES_TIPO', '" & gsEmpresa & "', '" & gsAnio & "', '" & Cuenta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "Seleccione un registro", vbInformation
        Set rsArreglo = Nothing
        Exit Function
    End If
    If rsArreglo(0).Value > 0 Then VerificaSiEsCuentasTipo = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing

End Function

Private Function VerificaMoneCtaCte(cCuentaConta As String, cMone As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    VerificaMoneCtaCte = False
    
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    sqlSp = "spCn_GrabaCuentaBanco 'VERIFMONE', '" & gsEmpresa & "','','','" & cCuentaConta & "', '" & cMone & "','','','',0,'','','',''"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
       Set rsArreglo = Nothing
       Exit Function
    End If
    
    If rsArreglo.RecordCount > 0 Then VerificaMoneCtaCte = True
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function VerificaCuentasMvtos(cCuentaConta As String) As Boolean
    
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp As String
    
    VerificaCuentasMvtos = False
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentas 'SEL_MVTOS', '" & gsEmpresa & "', '" & gsAnio & "', '" & cCuentaConta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
       Set rsArreglo = Nothing
       Exit Function
    End If
    
    If rsArreglo(0).Value > 0 Then VerificaCuentasMvtos = True
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function VerificaMovProvision(cCuentaConta As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp As String
    
    VerificaMovProvision = False
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentas 'SEL_PROV_CTA', '" & gsEmpresa & "', '" & gsAnio & "', '" & cCuentaConta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
       Set rsArreglo = Nothing
       Exit Function
    End If
    If rsArreglo.RecordCount > 0 Then VerificaMovProvision = True
    
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function EliminarCuenta(Cuenta As String)
    Dim respuesta As String, Pos As Integer
    
    If Cuenta <> "" Then
'        If VerificaCuentasHijas(Cuenta) = True Then
'            Mensajes "La cuenta " & Cuenta & " contiene una subcuenta y no puede eliminarse. Elimine subcuentas primero...", vbInformation
'            Exit Function
'        End If
'
'        If VerificaCuentasMvtos(Cuenta) = True Then
'            Mensajes "Se han registrado movimientos con esta cuenta " & Cuenta & ". Elimine los movimientos primero...", vbInformation
'            Exit Function
'        End If
        
        Dim clsMante As clsMantoTablas
        Set clsMante = New clsMantoTablas
        Call CargaArregloMnt
        lArrMnt(0) = "ELIMINAR"
        lArrMnt(3) = Cuenta
        clsMante.InicializaClase
        clsMante.BeginTrans
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCuentas", lArrMnt(), True) = False Then
            Mensajes "No se pudo eliminar esta la cuenta " & Cuenta, vbInformation
            Screen.MousePointer = vbDefault
            
            clsMante.CancelTrans
            clsMante.FinalizaClase
            
            Set clsMante = Nothing
            Exit Function
        End If
        clsMante.CommitTrans
        clsMante.FinalizaClase
        Mensajes "Registro ha sido eliminado", vbInformation
        Set clsMante = Nothing
    End If
End Function

Private Sub Borrar()
      Dim Eliminar As Boolean
      Dim Pos As Integer
      Dim i As Integer
      
      If VerificaSiEsCuentasTipo(tdbgCuentas.Columns(0).Value) = True Then
         Mensajes "La cuenta " & tdbgCuentas.Columns(0).Value & " es una cuenta que pertenece a un asiento tipo. " & Salto(1) & "Elimine primero el asiento tipo...", vbInformation
         Exit Sub
      End If
      
      If VerificaSiEsCuentasHija(tdbgCuentas.Columns(0).Value) = True Then
         Mensajes "La cuenta " & tdbgCuentas.Columns(0).Value & " es una cuenta de destino no se puede eliminar. Elimine subcuentas primero...", vbInformation
         Exit Sub
      End If
      
'      If VerificaCuentasHijas(tdbgCuentas.Columns(0).Value) = True Then
'         Mensajes "La cuenta " & tdbgCuentas.Columns(0).Value & " contiene una subcuenta y no puede eliminarse. Elimine subcuentas primero...", vbInformation
'         Exit Sub
'      End If
      
      If VerificaCuentasMvtos(tdbgCuentas.Columns(0).Value) = True Then
         Mensajes "Se han registrado movimientos con esta cuenta " & tdbgCuentas.Columns(0).Value & ". Elimine los movimientos primero...", vbInformation
         Exit Sub
      End If
      
      If Not lrsPlanCta Is Nothing Then
          If lrsPlanCta.State = adStateOpen Then
              If MsgBox("Desea eliminar definitivamente las cuentas seleccionadas", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro") = vbYes Then
                    Screen.MousePointer = vbHourglass
                    Pos = lrsPlanCta.AbsolutePosition
                    
                    If lrsPlanCta.RecordCount = 1 And tdbgCuentas.IsSelected(1) >= 0 Then
                        Eliminar = EliminarCuenta(tdbgCuentas.Columns(0).Value)
                        DoEvents
                    Else
                    
                        For i = 0 To tdbgCuentas.SelBookmarks.Count - 1
                            tdbgCuentas.Bookmark = tdbgCuentas.SelBookmarks(i)
                            Eliminar = EliminarCuenta(tdbgCuentas.Columns(0).Value)
                            DoEvents
                        Next i
                    
                    End If
                    
                    Call CargaTabla
                    
                    If Not lrsPlanCta Is Nothing Then
                        If Pos <= lrsPlanCta.RecordCount Then
                            On Error Resume Next
                            tdbgCuentas.Bookmark = Pos
                        End If
                    End If
                    
                    Screen.MousePointer = vbDefault
              
              End If
          Else
                Mensajes "Seleccione la cuenta a eliminar", vbOKOnly + vbInformation
          End If
      Else
        Mensajes "Seleccione la cuenta a eliminar", vbOKOnly + vbInformation
      
      End If
      
      Screen.MousePointer = vbNormal
End Sub

Private Sub ManNuevo()
    Dim sSql As String
    Dim RstDetalle As ADODB.Recordset
    sstParamatros.Tab = 0
    If Me.sstPerfiles.Tab = 1 Then Exit Sub
    lTipoMnt = "INSERTAR"
    
    Call LimpiaTexto(Me)
    Call EstadoChecks(False)
    tdbtBalance.Text = ""
    tdbtDual.Text = ""
    tdbtResFunc.Text = ""
    tdbtResNatu.Text = ""
    tdbtFlujoEfectivo.Text = "" 'frt_efe
        
    'Call HabilitaControl(Me)
    tdbtCodigo.Enabled = True
    tdbtCodigo.ReadOnly = False
    tdbcEntidad.Enabled = True
    tdbtDescripcion.Enabled = True
    chkTitulo.Enabled = True
    fraConfig.Enabled = True
    fradatos.Enabled = True
    tdbtCtaDestino.Enabled = True
    tdbtCtaDestino.ReadOnly = False
    tdbtCtaDestino.Text = ""
    tdbtNombreDestino.Text = ""
    
    Call Activarcampos(True)
    
    OptAct.Value = False
    OptPas.Value = False
    chkNat.Value = 0
    chkFun.Value = 0
    'ChkCtacte.Value = 0
    
    chkDetraccion.Value = 0
    chkRetencion.Value = 0
    chkPercepcion.Value = 0
    
    ' *******************************************
    
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    sstPerfiles.TabEnabled(2) = False
    sstParamatros.TabEnabled(1) = False
    
    tdbnPorc = 100

    tdblDestinoAux.Clear
    tdbcMes.Bookmark = 0
    Me.tdbcEntidad.Bookmark = 0
    Me.tdbcOperaTC.Bookmark = 0
    Me.tdbcPatrimomio.Bookmark = 0
    sstParamatros.Tab = 0
    
    Me.cmdEditar.Enabled = True
    
    cmdInsertar.Enabled = True
    cmdEliminarDestino.Enabled = True
       
    
  Dim clDatos As clsMantoTablas
  Dim arrDatos() As Variant

  sSql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Lib_cTipoLibro = 'LD'"

  Set clDatos = New clsMantoTablas
  arrDatos = Array(sSql)
  Set RstDetalle = New ADODB.Recordset
  Set RstDetalle = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    If Not RstDetalle Is Nothing Then
        If RstDetalle.RecordCount = 0 Then
            EstadoLDOri = "1"
        Else
            EstadoLDDes = "8"
        End If
    Else
        EstadoLDOri = "1"
    End If

    DoEvents
    tdbcMes.BoundText = "01"
    tdbcMes.ReBind
      
    Call chkProvision_Click
      
    sstParamatros.TabEnabled(0) = True
    pSetFocus tdbtCodigo
    
    Call HabilitarCampos
    tdbtCodigo.MaxLength = gsDigitosCtaDetalle
    
End Sub

Private Sub VerDatos()
    sstParamatros.Tab = 0
    If Me.sstPerfiles.Tab = 1 Then Exit Sub
    Call CargaDatosRegistro
    If lRegElim = False Then
        lblMante = "VER REGISTRO"
        sstPerfiles.TabEnabled(1) = True
        sstPerfiles.TabEnabled(0) = False
        sstPerfiles.Tab = 1
        
        tbrOpciones.Buttons(1).Enabled = False  ' *** nuevo
        tbrOpciones.Buttons(2).Enabled = False  ' *** consultar
        tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(5).Enabled = False  ' *** MODIFICAR
        tbrOpciones.Buttons(7).Image = 8
        
        Me.cmdEditar.Enabled = True
        
        sstPerfiles.TabEnabled(2) = False
        
        lTipoMnt = "EDITAR"
        
        Call Activarcampos(True)
        
        tdbcEntidad.Enabled = False
        tdbtDescripcion.Enabled = False
        chkTitulo.Enabled = False
        fraConfig.Enabled = False
        fradatos.Enabled = False
        tdbtCtaDestino.Enabled = False
        tdbtCtaDestino.ReadOnly = True
        tdbtCtaDestino.Text = ""
        tdbtNombreDestino.Text = ""
        
        DoEvents
        tdbcMes.BoundText = "01"
        tdbcMes.ReBind
        
        Call chkProvision_Click
        'Call AseguraControl(Me, True)
        
        'Call Activarcampos(False)
    Else
        lRegElim = False
    End If
    
    cmdInsertar.Enabled = False
    cmdEliminarDestino.Enabled = False
    sstParamatros.TabEnabled(0) = True
    
    'Call HabilitarCampos
End Sub

Private Sub Activarcampos(Opcion As Boolean)
        Me.tdbtBalance.Enabled = Opcion
        Me.OptAct.Enabled = Opcion
        Me.OptPas.Enabled = Opcion
        tdbtDual.Enabled = Opcion
        tdbtResFunc.Enabled = Opcion
        tdbtResNatu.Enabled = Opcion
        tdbtFlujoEfectivo.Enabled = Opcion 'frt_efe
        Me.chkCentroCosto.Enabled = Opcion
        Me.chkDetraccion.Enabled = Opcion
        Me.chkDocumento.Enabled = Opcion
        Me.ChkNCND.Enabled = Opcion
        Me.chkPercepcion.Enabled = Opcion
        Me.chkProvision.Enabled = Opcion
        Me.chkRetencion.Enabled = Opcion
        Me.chkTitulo.Enabled = Opcion
        
        Me.chkFun.Enabled = Opcion
        Me.chkNat.Enabled = Opcion
        
        tdbcPatrimomio.Enabled = Opcion
        Me.tdbtPla_cCuenta39.Enabled = Opcion
        If Opcion = True Then
            ActivarControl tdbcPatrimomio, Opcion
        End If
End Sub

Private Sub Editar()
    sstParamatros.Tab = 0
    'Me.chkFun.Enabled = False
    'Me.chkNat.Enabled = False

    OptAct.Value = False
    OptPas.Value = False
    chkNat.Value = 0
    chkFun.Value = 0
    
        'Me.tdbtBalance.Enabled = True
        'Me.OptAct.Enabled = True
        'Me.OptPas.Enabled = True
        'tdbtDual.Enabled = True
    
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        'If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        
        'Call HabilitaControl(Me)
        tdbcEntidad.Enabled = True
        tdbtDescripcion.Enabled = True
        
        chkTitulo.Enabled = True
        fraConfig.Enabled = True
        fradatos.Enabled = True
        tdbtCtaDestino.Enabled = True
        tdbtCtaDestino.ReadOnly = False
        tdbtCtaDestino.Text = ""
        tdbtNombreDestino.Text = ""
        
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
        sstPerfiles.TabEnabled(2) = False
    End If
    Call Activarcampos(True)
    DoEvents
    Call CargaDatosRegistro
    
    If Len(tdbtCodigo) = gsDigitosCtaDetalle Then
'        CmbRPI.Clear
'        CmbRPI.AddItem "", 0
'        CmbRPI.AddItem "AFP", 1
'        CmbRPI.AddItem "ONP", 2
        ChkConsPDT601.Visible = True
'        CmbRPI.Visible = True
    Else
        ChkConsPDT601.Visible = False
'        CmbRPI.Visible = False
    End If
    If lRegElim = False Then
        Select Case Left(tdbtCodigo, 2)
               Case "60", "62", "63", "64", "65", "67", "68"
                    sstParamatros.TabEnabled(1) = True
               Case Else
                    If Left(tdbtCodigo, 1) = "9" Then
                       sstParamatros.TabEnabled(1) = True
                    Else
                       sstParamatros.TabEnabled(1) = False
                    End If
        End Select
        
        tdbcPatrimomio.Enabled = True
        ActivarControl tdbcPatrimomio, True
        
        Me.cmdEditar.Enabled = True

        pSetFocus tdbtDescripcion
        auxDestino = False
    Else
        lRegElim = False
    End If
    
    EstadoLDDes = "9"
    
    DoEvents
    tdbcMes.BoundText = "01"
    tdbcMes.ReBind
    
    Call chkProvision_Click
        
    If Left(tdbtCodigo.Text, 1) = "9" Then
        Me.lblCostoProduccion.Visible = True
        Me.chkCostoProduccion.Visible = True
    Else
        Me.lblCostoProduccion.Visible = False
        Me.chkCostoProduccion.Visible = False
    End If
    
    If Left(tdbtCodigo.Text, 3) = "711" Or Left(tdbtCodigo.Text, 3) = "715" Then
        Me.lblVariacionProduccion.Visible = True
        Me.chkVariacionProduccion.Visible = True
    Else
        Me.lblVariacionProduccion.Visible = False
        Me.chkVariacionProduccion.Visible = False
    End If
        
    If tdbtCodigo.Text = "79" Or tdbtCodigo.Text = "69" Then
        Me.lblCuentaCostoVenta.Visible = True
        Me.chkCuentaCostoVenta.Visible = True
    Else
        Me.lblCuentaCostoVenta.Visible = False
        Me.chkCuentaCostoVenta.Visible = False
    End If
        
    cmdInsertar.Enabled = True
    cmdEliminarDestino.Enabled = True
    
    sstParamatros.TabEnabled(0) = True
    
    Call HabilitarCampos
    Call Habilita39(True)
    
End Sub

Private Function validarDatos() As Boolean
    Dim respuesta As String
    
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If TextoLleno2(Me.tdbtCodigo, "Codigo") = False Then Exit Function
    If TextoLleno2(Me.tdbtDescripcion, "Descripcion") = False Then Exit Function
    
'    'Validar la cuenta si tiene mov. y se cambia a Titulo
'    If chkTitulo.Value = 1 Then
'        If VerificaCuentasMvtos(Me.tdbtCodigo) = True Then
'           Mensajes "Esta cuenta tiene movimientos. Elimine los movimientos contables con esta cuenta...", vbInformation
'           Exit Function
'        End If
'    End If
'
'    'Valida si cuenta tiene mov. de provisión. No debe permitir modificar el check de provisión
'    If chkProvision.Value = 0 Then
'        If VerificaMovProvision(Me.tdbtCodigo) = True Then
'           Mensajes "Esta cuenta tiene movimientos de provisión. Verifique los movimientos contables con esta cuenta...", vbInformation
'           Exit Function
'        End If
'    End If
    
    'Validar Cuentas Destino
    If sstParamatros.TabEnabled(1) Then
        If tdblDestinoAux.ListCount > 0 Then
           If VerificarPorcDestino = False Then Exit Function
        Else
           ' *** Verificar q se ingrese destino si se requiere
           If chkTitulo = 0 And EsCtaDestino Then
              ' *** PREGUNTAR SI SE QUIERE INGRESAR DESTINO
              respuesta = MsgBox("Desea crear Destino a la cuenta seleccionada", vbYesNo + vbQuestion, "Confirmar Crear Destino")
              If respuesta = vbYes Then
                 sstParamatros.Tab = 1
                 pSetFocus tdbtCtaDestino
                 Exit Function
              End If
           End If
        End If
    End If
    
    ' Valida ingreso de Conf de Reportes
    If Len(CE(tdbtBalance.Text)) > 0 Then
        If ExisteCodigoRep("BGE", tdbtBalance) = False Then
            Mensajes "Codigo de Conasev BAL. GENERAL no registrado. Verifique...", vbInformation
            sstParamatros.Tab = 1
            pSetFocus tdbtBalance
            Exit Function
        End If
    End If
    
    If Len(CE(tdbtDual.Text)) > 0 Then
        If ExisteCodigoRep("BGE", tdbtDual) = False Then
            Mensajes "Codigo de Conasev BAL. GENERAL DUAL no registrado. Verifique...", vbInformation
            sstParamatros.Tab = 1
            pSetFocus tdbtDual
            Exit Function
        End If
    End If
    
    If Len(CE(tdbtResFunc.Text)) > 0 Then
        If ExisteCodigoRep("FUN", tdbtResFunc) = False Then
            Mensajes "Codigo de Conasev FUNCION no registrado. Verifique...", vbInformation
            sstParamatros.Tab = 1
            pSetFocus tdbtResFunc
            Exit Function
        End If
    End If
    
    If Len(CE(tdbtResNatu.Text)) > 0 Then
        If ExisteCodigoRep("NAT", tdbtResNatu) = False Then
            Mensajes "Codigo de Conasev NATURALEZA no registrado. Verifique...", vbInformation
            pSetFocus tdbtResNatu
            Exit Function
        End If
    End If
    
'    If Len(CE(tdbtFlujoEfectivo.Text)) > 0 Then 'frt_efe
'        If ExisteCodigoRep("EFE", tdbtFlujoEfectivo) = False Then
'            Mensajes "Codigo de FLUJO DE EFECTIVO no registrado. Verifique...", vbInformation
'            sstParamatros.Tab = 1
'            pSetFocus tdbtFlujoEfectivo
'            Exit Function
'        End If
'    End If
    
    ' ***
    validarDatos = True
End Function

Private Function BuscaCodigosOPChecks() As String
    BuscaCodigosOPChecks = ""
    
    If chkCierreVarias(0).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(0).Tag: Exit Function   'Cuenta de Utilidad
    If chkCierreVarias(1).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(1).Tag: Exit Function   'Cuenta de Perdida
    If chkCierreVarias(2).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(2).Tag: Exit Function   'Cuenta de Remuneraciones y Particip. por Pagar
    If chkCierreVarias(3).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(3).Tag: Exit Function   'Cuenta de Tributos por Pagar
    If chkCierreVarias(4).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(4).Tag: Exit Function   'Cuenta de variación de existencias
    If chkCierreVarias(5).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreVarias(5).Tag: Exit Function   'Cuenta de Reserva Legal

    If chkCierreCargas(0).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCargas(0).Tag: Exit Function   'costos de servicios
    If chkCierreCargas(1).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCargas(1).Tag: Exit Function   'gasto de ventas
    If chkCierreCargas(2).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCargas(2).Tag: Exit Function   'gastos administrativos
    If chkCierreCargas(3).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCargas(3).Tag: Exit Function   'gastos financieros

    If chkCierreCta8(0).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(0).Tag: Exit Function   'Margen Comercial
    If chkCierreCta8(1).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(1).Tag: Exit Function   'Valor Agregado
    If chkCierreCta8(2).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(2).Tag: Exit Function   'Excedente Bruto de Explotación
    If chkCierreCta8(3).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(3).Tag: Exit Function   'Resultado de Explotación
    If chkCierreCta8(4).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(4).Tag: Exit Function   'Resultado antes de Participaciones y Impuestos
    If chkCierreCta8(5).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(5).Tag: Exit Function   'Distribución Legal de Renta
    If chkCierreCta8(6).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(6).Tag: Exit Function   'Resultado del Ejercicio
    If chkCierreCta8(7).Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCta8(7).Tag: Exit Function   'Impuesto a la Renta
    If chkCierreCtaCnfImp.Value = vbChecked Then BuscaCodigosOPChecks = chkCierreCtaCnfImp.Tag: Exit Function   'Configuración de Impuesto
    
End Function

Private Sub AsignaCodigosOPChecks()
    
    chkCierreVarias(0).Tag = "037" 'Cuenta de Utilidad
    chkCierreVarias(1).Tag = "038" 'Cuenta de Perdida
    chkCierreVarias(2).Tag = "039" 'Cuenta de Remuneraciones y Particip. por Pagar
    chkCierreVarias(3).Tag = "040" 'Cuenta de Tributos por Pagar
    chkCierreVarias(4).Tag = "041" 'Cuenta de variación de existencias
    chkCierreVarias(5).Tag = "043" 'Cuenta de Reserva Legal

    chkCierreCargas(0).Tag = "042" 'costos de servicios
    chkCierreCargas(1).Tag = "044" 'gasto de ventas
    chkCierreCargas(2).Tag = "045" 'gastos administrativos
    chkCierreCargas(3).Tag = "046" 'gastos financieros

    chkCierreCta8(0).Tag = "029" 'Margen Comercial
    chkCierreCta8(1).Tag = "030" 'Valor Agregado
    chkCierreCta8(2).Tag = "031" 'Excedente Bruto de Explotación
    chkCierreCta8(3).Tag = "032" 'Resultado de Explotación
    chkCierreCta8(4).Tag = "033" 'Resultado antes de Participaciones y Impuestos
    chkCierreCta8(5).Tag = "034" 'Distribución Legal de Renta
    chkCierreCta8(6).Tag = "035" 'Resultado del Ejercicio
    chkCierreCta8(7).Tag = "036" 'Impuesto a la Renta
    chkCierreCtaCnfImp.Tag = "099" 'Configuración de Impuesto
    
    'chkCierreCta8(0).Enabled = True
End Sub


Private Function ValidaCuentaCierre() As Boolean
    On Error GoTo serror
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim cadena As String
    Dim Cuenta As String
    Set clDatos = New clsMantoTablas
    Dim rsDescrip As ADODB.Recordset
    Set rsDescrip = New ADODB.Recordset
    Dim CodOP As String
    
    CodOP = BuscaCodigosOPChecks()
        
    sqlSp = "spCn_ConsultaCuentas 'BUSCA_CTACIERRE', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbtCodigo.Text & "','" & CodOP & "'"
    arrDatos = Array(sqlSp)
    
    Set rsDescrip = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    cadena = ""
    
    If Not rsDescrip Is Nothing Then
        If Not (rsDescrip.EOF And rsDescrip.BOF) Then
            'cadena = cadena & rsDescrip.AbsolutePosition & ") " & CE(rsDescrip!COP_CDESCRIPCION) & Salto(1)
            cadena = CE(rsDescrip!COP_CDESCRIPCION)
            Cuenta = CE(rsDescrip!cod_cvalorparam)
            rsDescrip.MoveNext
        End If
    End If
    
    Call CerrarRecordSet(rsDescrip)
    
    If cadena = "" Then
        ValidaCuentaCierre = True
    Else
        Mensajes "La cuenta " & Cuenta & " ya fue configurada como:" & Salto(2) & cadena
        ValidaCuentaCierre = False
    End If
    
    Exit Function
serror:
    Mensajes Err.Description
    ValidaCuentaCierre = False
End Function

Private Function Validar() As Boolean
    Validar = False
    
    If CE(tdbtCodigo.Text) <> "" Then
        If Len(tdbtCodigo.Text) = gsDigitosCtaDetalle And chkTitulo.Value = vbChecked Then
            Mensajes "La cuenta ingresada es de tipo detalle, se desactivará la opcion de titulo "
            chkTitulo.Value = vbUnchecked
            pSetFocus chkTitulo
            Exit Function
        End If
    
        If Len(tdbtCodigo.Text) <> gsDigitosCtaDetalle And chkTitulo.Value = vbUnchecked Then
            Mensajes "La longitud de la cuenta de detalle debe ser " & gsDigitosCtaDetalle & " por que no es de tipo titulo"
            pSetFocus tdbtCodigo
            Exit Function
        End If
        
        If Len(tdbtCodigo.Text) = gsDigitosCtaDetalle And chkTitulo.Value = vbChecked Then
            Mensajes "La longitud de la cuenta de Titulo no debe ser igual" & Salto(1) & "a la longitud de la cuenta de detalle " & gsDigitosCtaDetalle
            pSetFocus tdbtCodigo
            Exit Function
        End If
    End If

    If chkNoTit(9).Value = vbChecked And Left(tdbtCodigo.Text, 1) <> "1" Then
        Mensajes "La cuenta ingresada no es una cuenta por cobrar"
        pSetFocus tdbtCodigo
        Exit Function
    End If

    If chkNoTit(10).Value = vbChecked And Left(tdbtCodigo.Text, 1) <> "4" Then
        Mensajes "La cuenta ingresada no es una cuenta por pagar"
        pSetFocus tdbtCodigo
        Exit Function
    End If

    If Len(tdbtCodigo.Text) <> gsDigitosCtaDetalle And tdblDestinoAux.ListCount > 0 Then
        Mensajes "La longitud de la cuenta " & tdbtCodigo.Text & " no es la de una cuenta de tipo detalle (" & gsDigitosCtaDetalle & " Digitos)"
        pSetFocus tdbtCodigo
        Exit Function
    End If

    If chkTitulo.Value = vbChecked And tdblDestinoAux.ListCount > 0 Then
        Mensajes "La cuenta ingresada es una cuenta de tipo titulo, no debe tener cuentas de destino "
        pSetFocus tdbtCodigo
        Exit Function
    End If
    
    If ValidaCuentaCierre = False Then
        Exit Function
    End If
    
    Validar = True
End Function

Private Sub Grabar()
    
    If Validar = False Then Exit Sub
    
    
    Dim clsMante As clsMantoTablas
    Dim i As Integer, condicion As Boolean
    Dim entro As Boolean
    
    If validarDatos = False Then Exit Sub
    
    Set clsMante = New clsMantoTablas
   
    '-----------------------------------------------
    ' GRABA CUENTA PRINCIPAL
    Call CargaArregloMnt
    condicion = False
    entro = False
    
    clsMante.InicializaClase
    clsMante.BeginTrans

    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCuentas", lArrMnt(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...'spCn_GrabaCuentas'", vbInformation
        
        clsMante.CancelTrans
        clsMante.FinalizaClase
        
        Exit Sub
    End If
    
    '----------------------------------------------------
    ' SI NO TIENE DESTINO PREGUNTA SI DESEA ELIMINAR TODOS LOS DESTINO DE TODO EL AÑO DE ESTA CUENTA
    
    If tdblDestinoAux.ListCount = 0 Then
        If (Left(CE(tdbtCodigo.Text), 1) = "6" Or Left(CE(tdbtCodigo.Text), 1) = "9") Then
           Dim Result As VbMsgBoxResult
        
            If sstParamatros.TabEnabled(1) = True And chkTitulo.Value = vbUnchecked Then
                Result = MsgBox("La cuenta " & tdbtCodigo.Text & " no tiene cuentas de destino para este mes," & Salto(2) & "Desea eliminar las cuentas de destino de todo el año de esta cuenta", vbQuestion + vbYesNo)
            Else
                Result = vbYes
            End If
            
            If Result = vbYes Then
                lArrMnt(0) = "ELIM_DESTINOANUAL"
                                
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCuentas", lArrMnt(), False) = False Then
                    Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                    
                    clsMante.CancelTrans
                    clsMante.FinalizaClase
                    
                    Exit Sub
                End If
                entro = True
            
            End If
        End If
    End If
    
    '----------------------------------------------------
    ' GRABA LA CUENTA DE DESTINO DEL MES SELECCIONADO
    
    If sstParamatros.TabEnabled(1) And entro = False Then
       ' *** Grabando Destino
       If tdblDestinoAux.ListCount > 0 Then
           '--------------------------
           'ELIMINA LOS DESTINOS DE LA CUENTA SELECCIONADA EN EL MES SELECCIONADO
           Call CargaArregloDestino(0)
           lArrDestino(0) = "ELIMINAR_MES"
           lArrDestino(3) = tdbcMes.BoundText
           
           If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrDestino(), False) = False Then
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
    
                clsMante.CancelTrans
                clsMante.FinalizaClase
    
                Exit Sub
           End If
           '--------------------------
           'INSERT LOS DESTINOS DE LA CUENTA SELECCIONADA EN EL MES SELECCIONADO
           
           For i = 0 To tdblDestinoAux.ListCount - 1
               If i = tdblDestinoAux.ListCount - 1 Then condicion = True
               Call CargaArregloDestino(i)
               If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrDestino(), False) = False Then
                   Mensajes "El proceso no ha concluido. Verificar...", vbInformation

                   clsMante.CancelTrans
                   clsMante.FinalizaClase

                   Exit Sub
               End If
           Next
       End If
       
    '---------------------------
    ' REPLICA EL DESTINO DEL MES SELECCIONADO A LOS MESES SIGUIENTE
    
       Call CargaArregloDestino(i)
       lArrDestino(0) = "REPLICA_DESTMESES"
       lArrDestino(3) = tdbcMes.BoundText
       lArrDestino(4) = tdbtCodigo.Text
       
       
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaDistCuentas", lArrDestino(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            
            clsMante.CancelTrans
            clsMante.FinalizaClase
            
            Exit Sub
        End If
    End If

    '---------------------------
    ' GRABA CONFIGURACIN DE OPERACIONES O CONFIG DE LAS CUENTAS

    If GrabaConfiguracionCuentasOp(clsMante) = False Then
        clsMante.CancelTrans
        clsMante.FinalizaClase
        Exit Sub
    End If
    
    '---------------------------
    
    clsMante.CommitTrans
    clsMante.FinalizaClase

    Call Cancelar
    Call CargaTabla
    
    ' *** Buscar La cuenta creada y posicionarse alli
    Dim Valor As Integer
    Valor = BuscarCadRs(tdbtCodigo, lrsPlanCta, 2)
    If Valor = 0 Then
            If Not (lrsPlanCta.BOF And lrsPlanCta.EOF) Then lrsPlanCta.MoveFirst
    End If
    Call ValidarCuentaCostoVenta
    ' ***
    Mensajes "Los datos se grabaron con exito...", vbInformation
    tdbgCuentas.HighlightRowStyle = "HighlightRow"
    Exit Sub
'ErrorEjecucion:
'    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub Cancelar()
    If Me.lblMante = "VER REGISTRO" Then
       ' Call AseguraControl(Me, False)
    Else
        'Call HabilitaControl(Me)
    End If
    Call TabMantenimiento(False)
    sstPerfiles.TabEnabled(2) = True
    sstPerfiles.Tab = 0
    tdbtCodigo.MaxLength = 0
    pSetFocus tdbgCuentas
End Sub

Private Sub TabMantenimiento(Valor As Boolean, Optional bFlag As Boolean = False)
    
    sstPerfiles.TabEnabled(1) = Valor
    sstPerfiles.TabEnabled(0) = Not Valor
    If Valor = True Then sstPerfiles.Tab = 1
    If Valor = False Then sstPerfiles.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    
    If Not bFlag Then
        tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
    Else
        tbrOpciones.Buttons(3).Enabled = Not Valor      ' *** Grabar
    End If
    
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
    
    If Valor = True Then
        Me.tbrOpciones.Buttons(7).Image = 8
    Else
        Me.tbrOpciones.Buttons(7).Image = 7
    End If
    
    
End Sub

Private Sub tdbcEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbcMes_GotFocus()
    If lTipoMnt = "INSERTAR" Then
        tdbcMes.Locked = False
    Else
        tdbcMes.Locked = False
        If tdblDestinoAux.ListCount = 0 Then Exit Sub
        If VerificarPorcDestino = False Then
            pSetFocus tdbtCodigo
        End If
    End If
    
End Sub

Private Sub tdbcMes_ItemChange()
    Call LlenaDestinoxMes(tdbcMes.BoundText)

End Sub

Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbcOperaTC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{TAB}"
End Sub

Private Sub tdbgCuentas_GotFocus()
    tdbgCuentas.HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgCuentas_HeadClick(ByVal ColIndex As Integer)
If Not lrsPlanCta Is Nothing Then
    If lrsPlanCta.RecordCount > 0 Then

        lrsPlanCta.Sort = tdbgCuentas.Columns(ColIndex).DataField
        tdbgCuentas.DataSource = lrsPlanCta
    End If
End If
End Sub

Private Sub tdbgCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Editar
End If
End Sub

Private Sub tdbgCuentas_LostFocus()
    tdbgCuentas.HighlightRowStyle = ""
End Sub

Private Sub tdbtBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtBalance.Enabled Then Call LlamaBuscar(frmBuscador, Me.tdbtBalance.Name, Control, "Balance", Me, gsPeriodo)
End Sub

Private Sub tdbtBalance_LostFocus()
    If sstPerfiles.Tab = 1 Then
        If Len(Trim(tdbtBalance.Text)) > 0 Then
            If ExisteCodigoRep("BGE", tdbtBalance) = False Then
                Mensajes "Codigo Reporte no registrado. Verifique...", vbInformation
                pSetFocus tdbtBalance
                tdbtBalance.Text = ""
            End If
        End If
    End If
End Sub

Private Sub tdbtCodGanancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtCodGanancia.ReadOnly = False Then Call LlamaBuscar(frmBuscador, Me.tdbtCodGanancia.Name, Control, "CuentasN", Me, gsPeriodo)
End Sub

Private Sub tdbtFlujoEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtFlujoEfectivo.Enabled Then Call LlamaBuscar(frmBuscador, Me.tdbtFlujoEfectivo.Name, Control, "FlujoEfectivo", Me, gsPeriodo)
End Sub

Private Sub tdbtFlujoEfectivo_LostFocus()
    If sstPerfiles.Tab = 1 Then
        If Len(Trim(tdbtFlujoEfectivo.Text)) > 0 Then
            If ExisteCodigoRep("EFE", tdbtFlujoEfectivo) = False Then
                Mensajes "Codigo Reporte no registrado. Verifique...", vbInformation
                pSetFocus tdbtFlujoEfectivo
                tdbtFlujoEfectivo.Text = ""
            End If
        End If
    End If
End Sub

Private Sub tdbtPla_cCuenta39_GotFocus()
'    ' *** Verificar q sea tipo de cuenta con destino
'    If chkTitulo = 0 And EsCtaDestino Then
'        tdbtPla_cCuenta39.ReadOnly = False
'    Else
'        tdbtPla_cCuenta39.ReadOnly = True
'    End If
'    ' ***
End Sub

Private Sub tdbtPla_cCuenta39_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If chkTitulo = 0 Then
            Call LlamaBuscar(frmBuscador, Me.tdbtPla_cCuenta39.Name, Control, "Cuentas39", Me, gsPeriodo, tdbtPla_cCuenta39.Text)
        Else
            Mensajes "Solo las cuentas de tipo detalle pueden tener cuentas de destino"
        End If
    End If
End Sub

Private Sub tdbtPla_cCuenta39_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
'    If sstParamatros.Tab = 1 Then
'        If Left(tdbtCodigo.Text, 1) = Left(tdbtCtaDestino.Text, 1) And EsCtaDestino Then
'            Mensajes "La cuenta no debe ser de la misma clase de la cuenta de origen " & Left(tdbtCodigo.Text, 1)
'            tdbtCtaDestino.Text = ""
'            tdbtNombreDestino.Text = ""
'            pSetFocus tdbtCtaDestino
'            Exit Sub
'        End If
        
    
        If tdbtPla_cCuenta39 <> "" And Me.tdbtPla_cCuenta39.Enabled = True Then
            'If Not fValidaCtaDestino(tdbtCtaDestino) Then tdbtCtaDestino = "": tdbtCtaDestino = "": Exit Sub
            tdbtPla_cCuenta39Nombre = ExisteCtaNoTitulo(tdbtPla_cCuenta39, "N")
            If tdbtPla_cCuenta39Nombre = "" Then pSetFocus tdbtPla_cCuenta39
        End If
'    End If
End Sub

Private Sub tdbtDescripcion_Change()
    If gsKey = 219 Then
        tdbtDescripcion = Replace(tdbtDescripcion, "'", "")
        tdbtDescripcion.SelStart = Len(tdbtDescripcion)
    End If
End Sub

Private Sub tdbtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
  gsKey = KeyCode
End Sub

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
  gsKey = KeyCode
End Sub

Private Sub tdbtDual_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = 112 And (CE(Left(tdbtCodigo, 2)) = "10" Or CE(Left(tdbtCodigo, 2)) = "12" Or CE(Left(tdbtCodigo, 2)) = "42") Then
'
'    Else
'        Mensajes "La cuenta debe pertenecer a la clase 10, 12 o 42 para ingresar datos en este campo", vbOKOnly + vbInformation
'    End If
    
    If KeyCode = 112 Then
        Call LlamaBuscar(frmBuscador, Me.tdbtDual.Name, Control, "Balance", Me, gsPeriodo)
    End If
End Sub

Private Sub tdbtDual_KeyPress(KeyAscii As Integer)
'    If CE(Left(tdbtCodigo, 2)) <> "10" And CE(Left(tdbtCodigo, 2)) <> "12" And CE(Left(tdbtCodigo, 2)) <> "42" Then
'        KeyAscii = 0
'        Exit Sub
'    End If

End Sub

Private Sub tdbtDual_LostFocus()
    If sstPerfiles.Tab = 1 Then
    If Len(Trim(tdbtDual.Text)) > 0 Then
        If ExisteCodigoRep("BGE", tdbtDual) = False Then
            Mensajes "Codigo Reporte no registrado. Verifique...", vbInformation
            pSetFocus tdbtDual
            tdbtDual.Text = ""
        End If
    End If
    End If
End Sub
Private Sub tdbtResFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtResFunc.Enabled Then Call LlamaBuscar(frmBuscador, Me.tdbtResFunc.Name, Control, "Funcion", Me, gsPeriodo)
End Sub
Private Sub tdbtResFunc_LostFocus()
    If sstPerfiles.Tab = 1 Then
    If Len(Trim(tdbtResFunc.Text)) > 0 Then
        If ExisteCodigoRep("FUN", tdbtResFunc) = False Then
            Mensajes "Codigo Reporte no registrado. Verifique...", vbInformation
            pSetFocus tdbtResFunc
            tdbtResFunc.Text = ""
        End If
    End If
    End If
    tdbtResFunc.Enabled = True
    
End Sub
Private Sub tdbtResNatu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtResNatu.Enabled Then Call LlamaBuscar(frmBuscador, Me.tdbtResNatu.Name, Control, "Naturaleza", Me, gsPeriodo)
End Sub
Private Sub tdbtResNatu_LostFocus()
    If sstPerfiles.Tab = 1 Then
    If Len(Trim(tdbtResNatu.Text)) > 0 Then
        If ExisteCodigoRep("NAT", tdbtResNatu) = False Then
            Mensajes "Codigo Reporte no registrado. Verifique...", vbInformation
            pSetFocus tdbtResNatu
            tdbtResNatu.Text = ""
        End If
    End If
    tdbtResNatu.Enabled = True
    End If
End Sub
Private Sub tdbtCodGanancia_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCodGanancia <> "" And Me.Enabled = True Then
        'pValidaCtaDestino
        tdbtDescGanancia = ExisteCtaNoTitulo(tdbtCodGanancia, "N")
        If tdbtDescGanancia = "" Then pSetFocus tdbtCodGanancia
    End If
End Sub
Private Sub tdbtCodigo_Change()
    If lblMante = "NUEVO REGISTRO" Then
       tdbcPatrimomio.Bookmark = 0
    End If
    
    Select Case Left(tdbtCodigo, 2)
        Case "60", "62", "63", "64", "65", "67", "68"
             sstParamatros.TabEnabled(1) = True
        Case Else
             If Left(tdbtCodigo, 1) = "9" Then
                sstParamatros.TabEnabled(1) = True
             Else
                sstParamatros.TabEnabled(1) = False
             End If
    End Select
    
    tdbcMes.Enabled = sstParamatros.TabEnabled(1)
    tdbcMes.Locked = Not sstParamatros.TabEnabled(1)
    
    ' Tipo de Cuenta
    If Left(tdbtCodigo, 1) >= "1" And Left(tdbtCodigo, 1) <= "5" Then
       OptAct.Enabled = True
       OptPas.Enabled = True
       chkNat.Value = vbUnchecked
       chkNat.Enabled = False
       chkFun.Value = vbUnchecked
       chkFun.Enabled = False
    Else
       OptAct.Enabled = False
       OptPas.Enabled = False
       OptAct.Value = False
       OptPas.Value = False
       chkNat.Enabled = True
       chkFun.Enabled = True
    End If
    
    '-----------------------------------------------
    ' OCULTA Y VISUALIZA FRAMES DE CONFIG DE OPERACIONES EN PLAN DE CUENTAS
    If Left(tdbtCodigo, 1) = "7" Then
        Call ActivarNoTituloCta456(vbUnchecked)
        Call ActivarNoTituloCta8(vbUnchecked)
        
    ElseIf Left(tdbtCodigo, 1) = "8" Then
        Call ActivarNoTituloCta7(vbUnchecked)
        Call ActivarNoTituloCta456(vbUnchecked)
    Else
        Call ActivarNoTituloCta7(vbUnchecked)
        Call ActivarNoTituloCta8(vbUnchecked)
        
    End If
    
    Call HabilitarCampos
End Sub
Private Sub tdbtCodigo_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbtCodigo.ReadOnly = True
    Else
        tdbtCodigo.ReadOnly = False
    End If
    ' ***
End Sub
Private Sub tdbtCodigo_LostFocus()

    ' *** Verificar q codigo exista
    If sstPerfiles.TabEnabled(1) = True And lTipoMnt = "INSERTAR" Then
        If ExisteCodigo(tdbtCodigo) = True Then
            Mensajes "Codigo ya existe. Verifique...", vbInformation
            pSetFocus tdbtCodigo
        End If
        ' *** Si es No Titulo; Verificar q tenga titulo a 2 digitos
        If Len(CE(tdbtCodigo)) > 2 Then
            If ExisteCodigo(Left(CE(tdbtCodigo), 2)) = False Then
                Mensajes "Cuenta titulo a 2 digitos aun no ha sido creada. Registrela primero.", vbInformation
                pSetFocus tdbtCodigo
            End If
        End If
        
        'Verifica si las cuentas de distribución
        Select Case Left(tdbtCodigo, 2)
               Case "60", "62", "63", "64", "65", "67", "68"
                    If Not fValidaCtaDestino(tdbtCodigo, False) Then
                    End If
               Case Else
                    If Left(tdbtCodigo, 1) = "9" Then
                       If Not fValidaCtaDestino(tdbtCodigo, False) Then
                       End If
                    End If
        End Select
        Call Habilita39(True)
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
    sqlSp = "spCn_ConsultaCuentas 'SEL_REG_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '" & Valor & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteCodigo = True
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Function ExisteCodigoRep(sTipo As String, sValor As String) As Boolean
    ' *** Verificar q codigo exista
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    ExisteCodigoRep = False
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaTipoPlantilla 'SEL_REG_CTA', '" & gsEmpresa & "', '" & sTipo & "', '" & sValor & "','','','','','','" & gsAnio & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteCodigoRep = True
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Private Sub tdbtCodigoBus_Change()
    Call FiltrarRecordSet
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(2) As String
    Dim i As Integer
    If lrsPlanCta Is Nothing Then Exit Sub
    cadena = ""
    If Trim(Me.tdbtCodigoBus) <> "" Then filtros(0) = "Pla_cCuentaContable like '" & tdbtCodigoBus & "*'"
    If Trim(Me.tdbtDescripcionBus) <> "" Then filtros(1) = "Pla_cNombreCuenta like '*" & tdbtDescripcionBus & "*'"
    If Trim(Me.tdbtTituloBus) <> "" Then filtros(2) = "Pla_cTitulo like '*" & tdbtTituloBus & "*'"
    For i = 0 To 2
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    
    If Trim(cadena) <> "" Then
        lrsPlanCta.Filter = cadena
    Else
        lrsPlanCta.Filter = 0
    End If
End Sub

Private Sub tdbtCodPerdida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtCodPerdida.ReadOnly = False Then Call LlamaBuscar(frmBuscador, tdbtCodPerdida.Name, Control, "CuentasN", Me, gsPeriodo)
End Sub

Private Sub tdbtCodPerdida_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCodPerdida <> "" And Me.Enabled = True Then
        'If Not fValidaCtaDestino(tdbtCodPerdida, True) Then tdbtCodPerdida = "": tdbtDescPerdida = "": Exit Sub
        tdbtDescPerdida = ExisteCtaNoTitulo(tdbtCodPerdida, "N")
        If tdbtDescPerdida = "" Then pSetFocus tdbtCodPerdida
    End If
End Sub
Private Sub tdbtCodRedondeoG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtCodRedondeoG.ReadOnly = False Then Call LlamaBuscar(frmBuscador, tdbtCodRedondeoG.Name, Control, "CuentasN", Me, gsPeriodo)
End Sub
Private Sub tdbtCodRedondeoG_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCodRedondeoG <> "" And Me.Enabled = True Then
       tdbtDescRedondeoG = ExisteCtaNoTitulo(tdbtCodRedondeoG, "N")
       If tdbtDescRedondeoG = "" Then pSetFocus tdbtCodRedondeoG
    End If
End Sub
Private Sub tdbtCodRedondeoP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And tdbtCodRedondeoP.ReadOnly = False Then Call LlamaBuscar(frmBuscador, tdbtCodRedondeoP.Name, Control, "CuentasN", Me, gsPeriodo)
End Sub
Private Sub tdbtCodRedondeoP_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCodRedondeoP <> "" And Me.Enabled = True Then
       tdbtDescRedondeoP = ExisteCtaNoTitulo(tdbtCodRedondeoP, "N")
       If tdbtDescRedondeoP = "" Then pSetFocus tdbtCodRedondeoP
    End If
End Sub

Private Sub tdbtCtaDestino_GotFocus()
    ' *** Verificar q sea tipo de cuenta con destino
    If chkTitulo = 0 And EsCtaDestino Then
        tdbtCtaDestino.ReadOnly = False
    Else
        tdbtCtaDestino.ReadOnly = True
    End If
    ' ***
End Sub

Private Sub tdbtCtaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        If chkTitulo = 0 And EsCtaDestino Then
            Call LlamaBuscar(frmBuscador, Me.tdbtCtaDestino.Name, Control, "CuentasN", Me, gsPeriodo, tdbtCtaDestino.Text)
        Else
            Mensajes "Solo las cuentas de tipo detalle pueden tener cuentas de destino"
        End If
    End If
End Sub

Private Sub tdbtCtaDestino_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If sstParamatros.Tab = 1 Then
'        If Left(tdbtCodigo.Text, 1) = Left(tdbtCtaDestino.Text, 1) And EsCtaDestino Then
'            Mensajes "La cuenta no debe ser de la misma clase de la cuenta de origen " & Left(tdbtCodigo.Text, 1)
'            tdbtCtaDestino.Text = ""
'            tdbtNombreDestino.Text = ""
'            pSetFocus tdbtCtaDestino
'            Exit Sub
'        End If
        
    
        If tdbtCtaDestino <> "" And Me.tdbtCtaDestino.Enabled = True Then
            'If Not fValidaCtaDestino(tdbtCtaDestino) Then tdbtCtaDestino = "": tdbtCtaDestino = "": Exit Sub
            tdbtNombreDestino = ExisteCtaNoTitulo(tdbtCtaDestino, "N")
            If tdbtNombreDestino = "" Then pSetFocus tdbtCtaDestino
        End If
    End If
End Sub

Private Sub tdbtDescripcionBus_Change()
    If gsKey = 219 Then
        tdbtDescripcionBus = Replace(tdbtDescripcionBus, "'", "")
        tdbtDescripcionBus.SelStart = Len(tdbtDescripcionBus)
    End If
    Call FiltrarRecordSet
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    
    ' *** Llenando el tipo de Entidad
    sqlcombos = "SELECT Ten_cTipoEntidad, Ten_cNombreEntidad From CNT_ENTIDAD "
    sqlcombos = sqlcombos + "WHERE Emp_cCodigo = '" & gsEmpresa & "' ORDER BY Ten_cNombreEntidad"
    LlenarComboAddItem Me.tdbcEntidad, sqlcombos, True
        
    ' *** Llenando Tipo de Operacion de Tipo de Cambio
    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo From TABLA "
    sqlcombos = sqlcombos + "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Tab_cTabla='026' ORDER BY Tab_cCodigo"
    LlenarComboAddItem tdbcOperaTC, sqlcombos, True

    ' *** Llena Patrimonio
    Dim arrA As New XArrayDB
    tdbcPatrimomio.Clear
    
    arrA.ReDim 0, 2, 0, 1
    arrA(0, 0) = ""
    arrA(0, 1) = "<NINGUNO>"
    
    arrA(1, 0) = "I"
    arrA(1, 1) = "INGRESO"
    
    arrA(2, 0) = "G"
    arrA(2, 1) = "GASTO"
    
    Set tdbcPatrimomio.Array = arrA
    tdbcPatrimomio.Bookmark = 0
    tdbcPatrimomio.ListField = "Column1"
    tdbcPatrimomio.BoundColumn = "Column0"
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Definir el mes en 01
    tdbcMes.Bookmark = 0
    tdbcPatrimomio.Bookmark = 0
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_ConsultaCuentas 'SEL_REG_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbgCuentas.Columns(0).Value & "'"
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
    tdbtCodigo = Trim(rsArreglo!Pla_cCuentaContable)
    tdbtDescripcion.Text = CE(rsArreglo!Pla_cNombreCuenta)
    If rsArreglo!Pla_cTitulo = "S" Then
        Me.chkTitulo.Value = 1
    Else
        Me.chkTitulo.Value = 0
    End If
    
    tdbcEntidad.BoundText = CE(rsArreglo!Ten_cTipoEntidad)
    tdbcOperaTC.BoundText = CE(rsArreglo!Pla_cOperaTC)
    
    '--------------- CHECKS DEL PLAN DE CTAS -------------'
    If IsNull(rsArreglo!Pla_cNCND) Then
        ChkNCND.Value = 0
    Else
        ChkNCND.Value = NE(rsArreglo!Pla_cNCND)
    End If

    If IsNull(rsArreglo!pla_cdetraccion) Then
        chkDetraccion.Value = 0
    Else
        chkDetraccion.Value = NE(rsArreglo!pla_cdetraccion)
    End If

    If IsNull(rsArreglo!pla_cretencion) Then
        chkRetencion.Value = 0
    Else
        chkRetencion.Value = NE(rsArreglo!pla_cretencion)
    End If

    If IsNull(rsArreglo!pla_cpercepcion) Then
        chkPercepcion.Value = 0
    Else
        chkPercepcion.Value = NE(rsArreglo!pla_cpercepcion)
    End If
    
    '-----------------------------------------------------'
    
    chkCentroCosto.Value = CE(rsArreglo!Pla_cCentroCosto)
    chkProvision.Value = IIf(IsNull(rsArreglo!Pla_cProvision) = True, 0, rsArreglo!Pla_cProvision)
    chkDocumento.Value = IIf(CE(rsArreglo!Pla_cDocumento) = "", "0", CE(rsArreglo!Pla_cDocumento))
    
'    TxtOrdCentrComp.Text = NE(rsArreglo!pla_OrdCentCom)
'    TxtOrdVta.Text = NE(rsArreglo!pla_OrdCentVta)

    If IsNull(rsArreglo!pla_cConsPDT) Then
        ChkConsPDT601.Value = 0
    Else
        ChkConsPDT601.Value = Val(rsArreglo!pla_cConsPDT)
    End If
    
    ' *** Nuevos parametros
    tdbcPatrimomio.BoundText = CE(rsArreglo!Pla_cCtaPresup)
    tdbtBalance = CE(rsArreglo!Pla_cCptoBG)
    tdbtDual = CE(rsArreglo!Pla_cCptoBGDual)
    
    tdbtResFunc = CE(rsArreglo!Pla_cCptoResFun)
    tdbtResNatu = CE(rsArreglo!Pla_cCptoResNat)
    'tdbtFlujoEfectivo = CE(rsArreglo!Pla_cCptoEFE) 'frt_efe
    
    Me.chkCuentaCostoVenta.Value = NE(rsArreglo!Pla_cCuentaCosVenta)
    Me.chkVariacionProduccion.Value = NE(rsArreglo!Pla_cVariacionProduccion)
    Me.chkCostoProduccion.Value = NE(rsArreglo!Pla_cCostoProduccion)
    ' *** Tipo de Cuenta
    Select Case CE(rsArreglo!Pla_cTipoCta)
        Case "A"                 'Activo
            OptAct.Value = True
        Case "P"                 'Pasivo
            OptPas.Value = True
        Case "R"                 'Naturaleza-Funcion
            chkNat.Value = 1
            chkFun.Value = 1
        Case "N"                 'Naturaleza
            chkNat.Value = 1
            chkFun.Value = 0
        Case "F"                 'Funcion
            chkFun.Value = 1
            chkNat.Value = 0
    End Select
    
    'ChkConsPDT601.Value = IIf(CE(Trim(rsArreglo!Pla_cCptoResFun)) = "", 0, Trim(rsArreglo!Pla_cCptoResFun))
    
'    If CE(Trim(rsArreglo!Pla_dRegimenIn)) = "1" Then
'        RegPI = "ONP"
'    ElseIf CE(Trim(rsArreglo!Pla_dRegimenIn)) = "2" Then
'        RegPI = "AFP"
'    Else
'        RegPI = ""
'    End If
    
    '-------------------------------------------------------------------
    Call LlenaDestinoxMes(tdbcMes.BoundText)
    '-------------------------------------------------------------------
    Call CargaConfiguracionCuentasOP
'    Me.tdbtPla_cCuenta39 = CE(rsArreglo!Pla_cCuenta39)
End Sub

Private Sub LlenaDestinoxMes(cMes As String)

    If CE(tdbtCodigo.Text) = "" Then
        tdblDestinoAux.Clear
        Exit Sub
    End If

    Dim sqlSp  As String
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    
    sqlSp = "spCn_ConsultaCuentas 'SEL_DES_CTADEST', '" & gsEmpresa & "', '" & gsAnio & "', '" & CE(tdbtCodigo.Text) & "','" & cMes & "'"
    arrDatos = Array(sqlSp)
    
    tdblDestinoAux.Clear
    
    Call CerrarRecordSet(rsArreglo)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then     ' *** Si no tiene destino
        tdblDestinoAux.Clear
        Set rsArreglo = Nothing
    Else                            ' *** Si tiene destino
        tdblDestinoAux.Clear
        Do While Not rsArreglo.EOF
            If CE(rsArreglo!Dis_cDestinoDebe) <> "" Then
                tdblDestinoAux.AddItem CE(rsArreglo!Dis_cDestinoDebe) & _
                "; " & CE(rsArreglo!Dis_cDestinoHaber) & "; " & CE(rsArreglo!Pla_cNombreCuenta) & " ;" & CE(rsArreglo!Dis_nPorcentaje)
                rsArreglo.MoveNext
            Else
                tdblDestinoAux.AddItem CE(rsArreglo!Dis_cDestinoDebe) & _
                "; " & CE(rsArreglo!Dis_cDestinoHaber) & "; " & CE(rsArreglo!Pla_cNombreCuentaHaber) & " ;" & CE(rsArreglo!Dis_nPorcentaje)
                rsArreglo.MoveNext
            End If
        Loop
    End If
    
    Set clDatos = Nothing
    Call CerrarRecordSet(rsArreglo)
End Sub

Private Sub EstadoChecks(Valor As Boolean)
    Dim nEstado As Integer
    If Valor = True Then
        nEstado = vbChecked
    Else
        nEstado = vbUnchecked
    End If
    
    chkTit(0).Value = nEstado 'chkTitBaseImpCompras
    chkTit(1).Value = nEstado 'chkTitBaseImpVentas
    chkTit(2).Value = nEstado 'chkTitGravColA
    chkTit(3).Value = nEstado 'chkTitGravColB
    chkTit(4).Value = nEstado 'chkTitGravColC
    chkTit(5).Value = nEstado 'chkTitIGV
    chkTit(6).Value = nEstado 'chkTitCtaCobrar
    chkTit(7).Value = nEstado 'chkTitCtaPagar
    chkTit(8).Value = nEstado 'chkTitCtaHonorarios
    chkTit(9).Value = nEstado 'chkTitProvTotalnestadoes

    chkNoTit(0).Value = nEstado 'chkNoTitCompHonorarios
    chkNoTit(1).Value = nEstado 'chkNoTitCompISC
    chkNoTit(2).Value = nEstado 'chkNoTitCompReintegro
    chkNoTit(3).Value = nEstado 'chkNoTitCompOtros
    chkNoTit(4).Value = nEstado 'chkNoTitVenExport
    chkNoTit(5).Value = nEstado 'chkNoTitVenBonifTxGrat
    
    chkNoTit(6).Value = nEstado 'chkNoTitVenBonifTxGrat
    chkNoTit(7).Value = nEstado 'chkNoTitVenBonifTxGrat
    
    chkNoTit(8).Value = nEstado 'chkNoTitLeasing
    chkNoTit(9).Value = nEstado 'chkNoLetCobrar
    chkNoTit(10).Value = nEstado 'chkNoLetPagar

    chkCierreVarias(0).Value = nEstado 'Cuenta de Utilidad
    chkCierreVarias(1).Value = nEstado 'Cuenta de Perdida
    chkCierreVarias(2).Value = nEstado 'Cuenta de Remuneraciones y Particip. por Pagar
    chkCierreVarias(3).Value = nEstado 'Cuenta de Tributos por Pagar
    chkCierreVarias(4).Value = nEstado 'Cuenta de variación de existencias
    chkCierreVarias(5).Value = nEstado 'Cuenta de Reserva Legal

    chkCierreCargas(0).Value = nEstado 'costos de servicios
    chkCierreCargas(1).Value = nEstado 'gasto de ventas
    chkCierreCargas(2).Value = nEstado 'gastos administrativos
    chkCierreCargas(3).Value = nEstado 'gastos financieros

    chkCierreCta8(0).Value = nEstado 'Margen Comercial
    chkCierreCta8(1).Value = nEstado 'nestado Agregado
    chkCierreCta8(2).Value = nEstado 'Excedente Bruto de Explotación
    chkCierreCta8(3).Value = nEstado 'Resultado de Explotación
    chkCierreCta8(4).Value = nEstado 'Resultado antes de Participaciones y Impuestos
    chkCierreCta8(5).Value = nEstado 'Distribución Legal de Renta
    chkCierreCta8(6).Value = nEstado 'Resultado del Ejercicio
    chkCierreCta8(7).Value = nEstado 'Impuesto a la Renta
    chkCierreCtaCnfImp.Value = nEstado 'Configuración de Impuesto

End Sub

Private Sub CargaConfiguracionCuentasOP()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp  As String
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentasConfig 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & CE(tdbgCuentas.Columns(0).Value) & "'"
    arrDatos = Array(sqlSp)
    
    Call CerrarRecordSet(rsArreglo)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then     ' *** Si no tiene datos
        
        Set rsArreglo = Nothing
    Else                            ' *** Si tiene datos
        Do While Not rsArreglo.EOF
            
            chkTit(0).Value = NE(rsArreglo!Pla_cTit_BaseImpCompras) 'chkTitBaseImpCompras
            chkTit(1).Value = NE(rsArreglo!Pla_cTit_BaseImpVentas) 'chkTitBaseImpVentas
            chkTit(2).Value = NE(rsArreglo!Pla_cTit_ColumnaA) 'chkTitGravColA
            chkTit(3).Value = NE(rsArreglo!Pla_cTit_ColumnaB) 'chkTitGravColB
            chkTit(4).Value = NE(rsArreglo!Pla_cTit_ColumnaC) 'chkTitGravColC
            chkTit(5).Value = NE(rsArreglo!Pla_cTit_IGV) 'chkTitIGV
            chkTit(6).Value = NE(rsArreglo!Pla_cTit_CtaCobrar) 'chkTitCtaCobrar
            chkTit(7).Value = NE(rsArreglo!Pla_cTit_CtaPagar) 'chkTitCtaPagar
            chkTit(8).Value = NE(rsArreglo!Pla_cTit_CtaPagarHonor) 'chkTitCtaHonorarios
            chkTit(9).Value = NE(rsArreglo!Pla_cTit_ProvTotalRepValores) 'chkTitProvTotalValores
        
            chkNoTit(0).Value = NE(rsArreglo!Pla_cNoTit_RegCompHonor) 'chkNoTitCompHonorarios
            chkNoTit(1).Value = NE(rsArreglo!Pla_cNoTit_RegCompISC) 'chkNoTitCompISC
            chkNoTit(2).Value = NE(rsArreglo!Pla_cNoTit_RegCompReint) 'chkNoTitCompReintegro
            chkNoTit(3).Value = NE(rsArreglo!Pla_cNoTit_RegCompOtros) 'chkNoTitCompOtros
            chkNoTit(4).Value = NE(rsArreglo!Pla_cNoTit_RegVenExp) 'chkNoTitVenExport
            chkNoTit(5).Value = NE(rsArreglo!Pla_cNoTit_RegVenBonif) 'chkNoTitVenBonifTxGrat
            
            chkNoTit(6).Value = NE(rsArreglo!Pla_cNoTit_Quinta) 'chkNoTitVenBonifTxGrat
            chkNoTit(7).Value = NE(rsArreglo!Pla_cNoTit_CuartaQuinta) 'chkNoTitVenBonifTxGrat
            
            chkNoTit(8).Value = NE(rsArreglo!Pla_cNoTit_Leasing) 'chkNoTitLeasing
            chkNoTit(9).Value = NE(rsArreglo!Pla_cNoTit_LetCobrar) 'chkNoLetCobrar
            chkNoTit(10).Value = NE(rsArreglo!Pla_cNoTit_LetPagar) 'chkNoLetPagar
        
            chkCierreVarias(0).Value = NE(rsArreglo!Pla_cNoTit456_Utilidad) 'Cuenta de Utilidad
            chkCierreVarias(1).Value = NE(rsArreglo!Pla_cNoTit456_Perdida) 'Cuenta de Perdida
            chkCierreVarias(2).Value = NE(rsArreglo!Pla_cNoTit456_Remun) 'Cuenta de Remuneraciones y Particip. por Pagar
            chkCierreVarias(3).Value = NE(rsArreglo!Pla_cNoTit456_Tributos) 'Cuenta de Tributos por Pagar
            chkCierreVarias(4).Value = NE(rsArreglo!Pla_cNoTit456_VarExis) 'Cuenta de variación de existencias
            chkCierreVarias(5).Value = NE(rsArreglo!Pla_cNoTit456_ResLegal) 'Cuenta de Reserva Legal
        
            chkCierreCargas(0).Value = NE(rsArreglo!Pla_cNoTit7_CostoServ) 'costos de servicios
            chkCierreCargas(1).Value = NE(rsArreglo!Pla_cNoTit7_GastoVentas) 'gasto de ventas
            chkCierreCargas(2).Value = NE(rsArreglo!Pla_cNoTit7_GastoAdm) 'gastos administrativos
            chkCierreCargas(3).Value = NE(rsArreglo!Pla_cNoTit7_GastoFinac) 'gastos financieros
        
            chkCierreCta8(0).Value = NE(rsArreglo!Pla_cNoTit8_MargComerc) 'Margen Comercial
            chkCierreCta8(1).Value = NE(rsArreglo!Pla_cNoTit8_ValAgreg) 'Valor Agregado
            chkCierreCta8(2).Value = NE(rsArreglo!Pla_cNoTit8_ExBrutoExplot) 'Excedente Bruto de Explotación
            chkCierreCta8(3).Value = NE(rsArreglo!Pla_cNoTit8_ResulExplot) 'Resultado de Explotación
            chkCierreCta8(4).Value = NE(rsArreglo!Pla_cNoTit8_ResulAntPartImp) 'Resultado antes de Participaciones y Impuestos
            chkCierreCta8(5).Value = NE(rsArreglo!Pla_cNoTit8_DistLegalRenta) 'Distribución Legal de Renta
            chkCierreCta8(6).Value = NE(rsArreglo!Pla_cNoTit8_ResultEjer) 'Resultado del Ejercicio
            chkCierreCta8(7).Value = NE(rsArreglo!Pla_cNoTit8_ImpRenta) 'Impuesto a la Renta
            chkCierreCtaCnfImp.Value = NE(rsArreglo!Pla_cConfigImp) 'Configuración de Impuesto
        
            rsArreglo.MoveNext
        Loop
    End If

    Set clDatos = Nothing
    Call CerrarRecordSet(rsArreglo)

End Sub

Private Function GrabaConfiguracionCuentasOp(ByRef oClase As clsMantoTablas) As Boolean
    GrabaConfiguracionCuentasOp = False
'    Dim clsMante As clsMantoTablas
'    Set clsMante = New clsMantoTablas

    On Local Error GoTo ErrorEjecucion
    
    Call CargaArregloConfOPCta
    
'    clsMante.InicializaClase
'    clsMante.BeginTrans
'
    
    If oClase.MantenimientoDeTablas(gsCadenaConexion, "spCn_ConsultaCuentasConfig", lArrMnt(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        
'        clsMante.CancelTrans
'        clsMante.FinalizaClase
        
        Exit Function
    End If
    
'    clsMante.CommitTrans
'    clsMante.FinalizaClase
'
'    Set clsMante = Nothing

    GrabaConfiguracionCuentasOp = True
    Exit Function
ErrorEjecucion:
    Mensajes Err.Description
End Function

Private Sub CargaArregloConfOPCta()
    ReDim lArrMnt(44) As Variant
    lArrMnt(0) = "INSERTAR"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = tdbtCodigo.Text

    lArrMnt(4) = chkTit(0).Value 'chkTitBaseImpCompras
    lArrMnt(5) = chkTit(1).Value 'chkTitBaseImpVentas
    lArrMnt(6) = chkTit(2).Value 'chkTitGravColA
    lArrMnt(7) = chkTit(3).Value 'chkTitGravColB
    lArrMnt(8) = chkTit(4).Value 'chkTitGravColC
    lArrMnt(9) = chkTit(5).Value 'chkTitIGV
    lArrMnt(10) = chkTit(6).Value 'chkTitCtaCobrar
    lArrMnt(11) = chkTit(7).Value 'chkTitCtaPagar
    lArrMnt(12) = chkTit(8).Value 'chkTitCtaHonorarios
    lArrMnt(13) = chkTit(9).Value 'chkTitProvTotalValores

    lArrMnt(14) = chkNoTit(0).Value 'chkNoTitCompHonorarios
    lArrMnt(15) = chkNoTit(1).Value 'chkNoTitCompISC
    lArrMnt(16) = chkNoTit(2).Value 'chkNoTitCompReintegro
    lArrMnt(17) = chkNoTit(3).Value 'chkNoTitCompOtros
    lArrMnt(18) = chkNoTit(4).Value 'chkNoTitVenExport
    lArrMnt(19) = chkNoTit(5).Value 'chkNoTitVenBonifTxGrat

    lArrMnt(20) = chkCierreVarias(0).Value 'Cuenta de Utilidad
    lArrMnt(21) = chkCierreVarias(1).Value 'Cuenta de Perdida
    lArrMnt(22) = chkCierreVarias(2).Value 'Cuenta de Remuneraciones y Particip. por Pagar
    lArrMnt(23) = chkCierreVarias(3).Value 'Cuenta de Tributos por Pagar
    lArrMnt(24) = chkCierreVarias(4).Value 'Cuenta de variación de existencias
    lArrMnt(25) = chkCierreVarias(5).Value 'Cuenta de Reserva Legal

    lArrMnt(26) = chkCierreCargas(0).Value 'costos de servicios
    lArrMnt(27) = chkCierreCargas(1).Value 'gasto de ventas
    lArrMnt(28) = chkCierreCargas(2).Value 'gastos administrativos
    lArrMnt(29) = chkCierreCargas(3).Value 'gastos financieros

    lArrMnt(30) = chkCierreCta8(0).Value 'Margen Comercial
    lArrMnt(31) = chkCierreCta8(1).Value 'Valor Agregado
    lArrMnt(32) = chkCierreCta8(2).Value 'Excedente Bruto de Explotación
    lArrMnt(33) = chkCierreCta8(3).Value 'Resultado de Explotación
    lArrMnt(34) = chkCierreCta8(4).Value 'Resultado antes de Participaciones y Impuestos
    lArrMnt(35) = chkCierreCta8(5).Value 'Distribución Legal de Renta
    lArrMnt(36) = chkCierreCta8(6).Value 'Resultado del Ejercicio
    lArrMnt(37) = chkCierreCta8(7).Value 'Impuesto a la Renta
    
    lArrMnt(38) = chkNoTit(6).Value 'Quinta
    lArrMnt(39) = chkNoTit(7).Value 'Cuarta quinta
    
    lArrMnt(40) = chkNoTit(8).Value 'leasing
    lArrMnt(41) = chkNoTit(9).Value 'cuenta de letra por cobrar
    lArrMnt(42) = chkNoTit(10).Value 'cuenta de letra por pagar
    
    lArrMnt(43) = gsUsuario
    
    lArrMnt(44) = chkCierreCtaCnfImp.Value 'Configuración de Impuesto
    
End Sub

Private Sub CargaArregloDestino(Numero As Integer)
    ReDim lArrDestino(9) As Variant
    lArrDestino(0) = lTipoMnt                               ' Accion
    lArrDestino(1) = gsEmpresa                              ' Empresa
    lArrDestino(2) = gsAnio                                 ' Anio
    tdblDestinoAux.Bookmark = Numero
    lArrDestino(3) = tdbcMes.BoundText   ' mes
    lArrDestino(4) = tdbtCodigo                             ' Cuenta
    lArrDestino(5) = Numero + 1                             ' Secuencia
    
    If NE(tdblDestinoAux.Columns(0).Value) <> 0 Then
        lArrDestino(6) = CE(tdblDestinoAux.Columns(0).Value)
        lArrDestino(7) = ""
    End If
    
    If NE(tdblDestinoAux.Columns(1).Value) <> 0 Then
        lArrDestino(6) = ""
        lArrDestino(7) = CE(tdblDestinoAux.Columns(1).Value)
        
    End If
    
    
    lArrDestino(8) = NE(tdblDestinoAux.Columns(3).Value)           ' Porcentaje
    lArrDestino(9) = gsUsuario                              ' Usuario
End Sub

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    Dim rs As New ADODB.Recordset
    
    ReDim lArrMnt(29) As Variant
'    ReDim lArrMnt(26) As Variant
    lArrMnt(0) = lTipoMnt                           ' Accion
    lArrMnt(1) = gsEmpresa                          ' Empresa
    lArrMnt(2) = gsAnio                             ' Año
    lArrMnt(3) = tdbtCodigo                         ' Cuenta
    lArrMnt(4) = CE(tdbtDescripcion)                ' Nombre Cuenta
    If chkTitulo.Value = 1 Then
        lArrMnt(5) = "S"
    Else
        lArrMnt(5) = "N"
    End If
    
    lArrMnt(6) = CE(tdbcEntidad.BoundText)          ' Entidad
    lArrMnt(7) = chkCentroCosto                     ' Centro Costo
    lArrMnt(8) = chkProvision                       ' Provision
    lArrMnt(9) = CE(tdbcOperaTC.BoundText)          ' Tipo Operacion TC
    lArrMnt(10) = chkDocumento.Value                ' Documento
    
    lArrMnt(11) = ""
    
    If OptAct.Value = True Then lArrMnt(11) = "A"   ' Tipo de Cuenta Activo
    If OptPas.Value = True Then lArrMnt(11) = "P"   ' Tipo de Cuenta Pasivo
    If chkNat.Value = 1 Then lArrMnt(11) = "N"      ' Tipo de Cuenta Naturaleza
    If chkFun.Value = 1 Then lArrMnt(11) = "F"      ' Tipo de Cuenta Funcion
    
    If chkNat.Value = 1 And chkFun.Value = 1 Then lArrMnt(11) = "R"
    
    lArrMnt(12) = tdbtBalance.Text                  ' Config Balance
    lArrMnt(13) = tdbtDual.Text                     ' Config Balance dual
    lArrMnt(14) = tdbtResFunc.Text                  ' Config Función
    lArrMnt(15) = tdbtResNatu.Text                  ' Config Naturaleza
    lArrMnt(16) = tdbcPatrimomio.Columns(0).Value   ' Patrimonio
    
    lArrMnt(17) = gsUsuario                         ' Usuario
    lArrMnt(18) = "A"                               ' Estado
    lArrMnt(19) = ChkNCND.Value                     ' Tiene NCND
    lArrMnt(20) = chkDetraccion.Value               ' Detraccion
    lArrMnt(21) = chkRetencion.Value                ' Retencion
    lArrMnt(22) = chkPercepcion.Value               ' Percepcion
    lArrMnt(23) = ChkConsPDT601.Value               ' considerar para PDt
'    lArrMnt(24) = Val(TxtOrdCentrComp.Text)        ' Orden de Centralizacion Compra
'    lArrMnt(25) = Val(TxtOrdVta.Text)              ' Orden de Centralizacion Ventas
'    lArrMnt(24) = tdbtPla_cCuenta39.Text           ' Cuenta39
    lArrMnt(24) = EstadoLDOri                       ' Estado Inicial de la cuenta 1
    lArrMnt(25) = EstadoLDDes                       ' Estado posterior al inicial 8 y 9
'    If Trim(CmbRPI.Text) = "" Then                 ' Indicador del Regimen Pensionario Independiente
'        lArrMnt(26) = "3"
'    ElseIf Trim(CmbRPI.Text) = "AFP" Then
'        lArrMnt(26) = "2"
'    ElseIf Trim(CmbRPI.Text) = "ONP" Then
'        lArrMnt(26) = "1"
'    End If
    lArrMnt(26) = CE(Me.chkCuentaCostoVenta.Value)
    lArrMnt(27) = CE(Me.chkVariacionProduccion.Value)
    lArrMnt(28) = CE(Me.chkCostoProduccion.Value)
    lArrMnt(29) = tdbtFlujoEfectivo.Text             ' Config Flujo Efectivo 'frt_efe
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
           Case "tdbtCtaDestino"   ' *** Caso de cliente
                tdbtCtaDestino = Trim(param0)
                Me.tdbtNombreDestino = Trim(param1)
                Unload frmBuscador
                pSetFocus cmbTipo
           Case "tdbtCodGanancia"  ' *** Caso Cta Ganancia
                tdbtCodGanancia = Trim(param0)
                tdbtDescGanancia = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtCodGanancia
           Case "tdbtCodPerdida"   ' *** Caso Cta Perdida
                tdbtCodPerdida = Trim(param0)
                tdbtDescPerdida = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtCodPerdida
           Case "tdbtCodRedondeoG"   ' *** Caso Cta Redondeo Ganancia
                tdbtCodRedondeoG = Trim(param0)
                tdbtDescRedondeoG = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtCodRedondeoG
           Case "tdbtCodRedondeoP"   ' *** Caso Cta Redondeo Perdida
                tdbtCodRedondeoP = Trim(param0)
                tdbtDescRedondeoP = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtCodRedondeoP
           Case "tdbtBalance"
                tdbtBalance = Trim(param0)
                Unload frmBuscador
                pSetFocus tdbtBalance
           Case "tdbtDual"
                tdbtDual = Trim(param0)
                Unload frmBuscador
                pSetFocus tdbtDual
           Case "tdbtResFunc"
                tdbtResFunc = Trim(param0)
                Unload frmBuscador
                pSetFocus tdbtResFunc
           Case "tdbtResNatu"
                tdbtResNatu = Trim(param0)
                Unload frmBuscador
                pSetFocus tdbtResNatu
           Case "tdbtPla_cCuenta39"   ' *** Caso Cta 39
                tdbtPla_cCuenta39 = Trim(param0)
                tdbtPla_cCuenta39Nombre = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtPla_cCuenta39
           Case "tdbtFlujoEfectivo" 'frt_efe
                tdbtFlujoEfectivo = Trim(param0)
                Unload frmBuscador
                pSetFocus tdbtFlujoEfectivo
    End Select
End Sub

Private Sub tdbtTituloBus_Change()
    Call FiltrarRecordSet
End Sub

Private Function fValidaCtaDestino(oText As TDBText, bFlag As Boolean) As Boolean

    Dim ObjCons As clsMantoTablas
    Dim rs As ADODB.Recordset
    Dim arr() As Variant
    Dim sqlSp As String
    
    sqlSp = "spCn_GrabaDistCuentas 'BUSCA_CTA_DEST', '" & gsEmpresa & "', '" & gsAnio & "','','" & oText.Text & "'"
    Set ObjCons = New clsMantoTablas
    arr = Array(sqlSp)
    
    'Call CerrarRecordSet(rsArreglo)
    
    Set rs = ObjCons.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
    
    If rs Is Nothing Then
        If bFlag Then
            MsgBox "NO HA DEFINIDO SUS CUENTAS DE GASTOS, INGRESE SUS CUENTAS DE DESTINO " + Chr(13) + _
                   "Y LUEGO ASIGNE UNA CUENTA DE DESTINO, PARA LA CUENTA DE PERDIDA X DIF. CAMB." + CE(rs.Fields("Pla_cCuentaContable").Value) + "' ", vbOKOnly + vbInformation, gsNombreModulo
        End If
        fValidaCtaDestino = False
        Exit Function
    End If

    fValidaCtaDestino = True

End Function

Private Sub tdbtTituloBus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 83 Or KeyAscii = 78 Or KeyAscii = 8 Or KeyAscii = 110 Or KeyAscii = 115 Then
        
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub ActivarTitulos(Valor As Integer)
    chkTit(0).Value = Valor 'chkTitBaseImpCompras
    chkTit(1).Value = Valor 'chkTitBaseImpVentas
    chkTit(2).Value = Valor 'chkTitGravColA
    chkTit(3).Value = Valor 'chkTitGravColB
    chkTit(4).Value = Valor 'chkTitGravColC
    chkTit(5).Value = Valor 'chkTitIGV
    chkTit(6).Value = Valor 'chkTitCtaCobrar
    chkTit(7).Value = Valor 'chkTitCtaPagar
    chkTit(8).Value = Valor 'chkTitCtaHonorarios
    chkTit(9).Value = Valor 'chkTitProvTotalValores
End Sub

Private Sub ActivarNoTitulos(Valor As Integer)
    chkNoTit(0).Value = Valor 'chkNoTitCompHonorarios
    chkNoTit(1).Value = Valor 'chkNoTitCompISC
    chkNoTit(2).Value = Valor 'chkNoTitCompReintegro
    chkNoTit(3).Value = Valor 'chkNoTitCompOtros
    chkNoTit(4).Value = Valor 'chkNoTitVenExport
    chkNoTit(5).Value = Valor 'chkNoTitVenBonifTxGrat
    
    chkNoTit(8).Value = Valor 'chkNoTitLeasing
    chkNoTit(9).Value = Valor 'chkNoLetCobrar
    chkNoTit(10).Value = Valor 'chkNoLetPagar
End Sub

Private Sub ActivarNoTituloCta7(Valor As Integer)
    chkCierreCargas(0).Value = Valor 'costos de servicios
    chkCierreCargas(1).Value = Valor 'gasto de ventas
    chkCierreCargas(2).Value = Valor 'gastos administrativos
    chkCierreCargas(3).Value = Valor 'gastos financieros
End Sub

Private Sub ActivarNoTituloCta456(Valor As Integer)
    chkCierreVarias(0).Value = Valor 'Cuenta de Utilidad
    chkCierreVarias(1).Value = Valor 'Cuenta de Perdida
    chkCierreVarias(2).Value = Valor 'Cuenta de Remuneraciones y Particip. por Pagar
    chkCierreVarias(3).Value = Valor 'Cuenta de Tributos por Pagar
    chkCierreVarias(4).Value = Valor 'Cuenta de variación de existencias
    chkCierreVarias(5).Value = Valor 'Cuenta de Reserva Legal
End Sub

Private Sub ActivarNoTituloCta8(Valor As Integer)
    chkCierreCta8(0).Value = Valor 'Margen Comercial
    chkCierreCta8(1).Value = Valor 'Valor Agregado
    chkCierreCta8(2).Value = Valor 'Excedente Bruto de Explotación
    chkCierreCta8(3).Value = Valor 'Resultado de Explotación
    chkCierreCta8(4).Value = Valor 'Resultado antes de Participaciones y Impuestos
    chkCierreCta8(5).Value = Valor 'Distribución Legal de Renta
    chkCierreCta8(6).Value = Valor 'Resultado del Ejercicio
    chkCierreCta8(7).Value = Valor 'Impuesto a la Renta
    chkCierreCtaCnfImp.Value = Valor 'Configuración de Impuesto
End Sub

Private Sub chkCierreCargas_Click(Index As Integer)
    On Error GoTo serror
    Dim entro As Boolean, i As Integer
    entro = False
    If chkCierreCargas(Index).Value = vbChecked Then
        entro = True
    End If

    Call ActivarTitulos(vbUnchecked)
    Call ActivarNoTitulos(vbUnchecked)
    Call ActivarNoTituloCta456(vbUnchecked)
    Call ActivarNoTituloCta8(vbUnchecked)
    
    For i = 0 To 3
        If i <> Index And chkCierreCargas(i).Value = vbChecked Then chkCierreCargas(i).Value = vbUnchecked
    Next i
    
    If entro Then
        chkCierreCargas(Index).Value = vbChecked
    End If
    
    Exit Sub
serror:
End Sub

Private Sub chkCierreVarias_Click(Index As Integer)
    On Error GoTo serror
    Dim entro As Boolean, i As Integer
    entro = False
    If chkCierreVarias(Index).Value = vbChecked Then
        entro = True
    End If

    Call ActivarTitulos(vbUnchecked)
    Call ActivarNoTitulos(vbUnchecked)
    Call ActivarNoTituloCta7(vbUnchecked)
    Call ActivarNoTituloCta8(vbUnchecked)
    
    For i = 0 To 5
        If i <> Index And chkCierreVarias(i).Value = vbChecked Then chkCierreVarias(i).Value = vbUnchecked
    Next i
    
    If entro Then
        chkCierreVarias(Index).Value = vbChecked
    End If
    
    Exit Sub
serror:
End Sub

Private Sub chkCierreCta8_Click(Index As Integer)
    On Error GoTo serror
    Dim entro As Boolean, i As Integer
    entro = False
    If chkCierreCta8(Index).Value = vbChecked Then
        entro = True
    End If

    Call ActivarTitulos(vbUnchecked)
    Call ActivarNoTitulos(vbUnchecked)
    Call ActivarNoTituloCta7(vbUnchecked)
    Call ActivarNoTituloCta456(vbUnchecked)
    
    For i = 0 To 7
        If i <> Index And chkCierreCta8(i).Value = vbChecked Then chkCierreCta8(i).Value = vbUnchecked
    Next i
    
    If entro Then
        chkCierreCta8(Index).Value = vbChecked
    End If
    
    Exit Sub
serror:
End Sub

Private Sub LimpiarTodoCheck()
    Call ActivarTitulos(vbUnchecked)
    Call ActivarNoTitulos(vbUnchecked)
    Call ActivarNoTituloCta7(vbUnchecked)
    Call ActivarNoTituloCta456(vbUnchecked)
    Call ActivarNoTituloCta8(vbUnchecked)
End Sub

Private Sub chkTit_Click(Index As Integer)
    On Error GoTo serror
    Dim entro As Boolean
    entro = False
    If chkTit(Index).Value = vbChecked Then
        entro = True
    End If
    
    Call ActivarNoTitulos(vbUnchecked)
    Call ActivarNoTituloCta7(vbUnchecked)
    Call ActivarNoTituloCta456(vbUnchecked)
    Call ActivarNoTituloCta8(vbUnchecked)
    
    If entro Then
        chkTit(Index).Value = vbChecked
    End If
    
    'If Index = 6 Or Index = 7 Then
    If chkTit(Index).Value = vbChecked And Len(tdbtCodigo.Text) <> 2 And chkTitulo.Value = vbChecked Then
        Mensajes "Solo debe ser asignada a una cuenta titulo de dos digitos"
        chkTit(Index).Value = vbUnchecked
    End If
    'End If
    
    Exit Sub
serror:
End Sub

Private Sub chkNoTit_Click(Index As Integer)
    On Error GoTo serror
    Dim entro As Boolean
    entro = False
    If chkNoTit(Index).Value = vbChecked Then
        entro = True
    End If

    Call ActivarTitulos(vbUnchecked)
    Call ActivarNoTituloCta7(vbUnchecked)
    Call ActivarNoTituloCta456(vbUnchecked)
    Call ActivarNoTituloCta8(vbUnchecked)
    
    If entro Then
        chkNoTit(Index).Value = vbChecked
    End If
    
    
    Exit Sub
serror:
End Sub

Sub Habilita39(Habilita As Boolean)
''*****HABILITA CAMPOS DE LA CUENTA 39*****
'    If Left(Me.tdbtCodigo.Text, 2) = "32" Then
'        Me.tdbtPla_cCuenta39.Enabled = Habilita
'        Me.Label13(11).Visible = Habilita
'        Me.tdbtPla_cCuenta39.Visible = Habilita
'        Me.tdbtPla_cCuenta39Nombre.Visible = Habilita
'    End If
End Sub

