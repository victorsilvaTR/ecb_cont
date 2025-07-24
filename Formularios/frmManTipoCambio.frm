VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tipo de Cambio"
   ClientHeight    =   5985
   ClientLeft      =   855
   ClientTop       =   2970
   ClientWidth     =   7965
   Icon            =   "frmManTipoCambio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7965
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5520
      Left            =   45
      TabIndex        =   26
      Top             =   405
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9737
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   706
      TabCaption(0)   =   "Moneda Extranjera"
      TabPicture(0)   =   "frmManTipoCambio.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Adicionales"
      TabPicture(1)   =   "frmManTipoCambio.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mensual"
      TabPicture(2)   =   "frmManTipoCambio.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Mantenimiento"
      TabPicture(3)   =   "frmManTipoCambio.frx":0F1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(0)"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4665
         Index           =   1
         Left            =   -74730
         TabIndex        =   42
         Top             =   585
         Width           =   7095
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   135
            TabIndex        =   43
            Top             =   1170
            Width           =   6825
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   3
               Left            =   1620
               TabIndex        =   12
               Tag             =   "enabled"
               Top             =   1485
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":0F3A
               Caption         =   "frmManTipoCambio.frx":0F5A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":0FC6
               Keys            =   "frmManTipoCambio.frx":0FE4
               Spin            =   "frmManTipoCambio.frx":1044
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   4
               Left            =   1620
               TabIndex        =   13
               Tag             =   "enabled"
               Top             =   1890
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":106C
               Caption         =   "frmManTipoCambio.frx":108C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":10F8
               Keys            =   "frmManTipoCambio.frx":1116
               Spin            =   "frmManTipoCambio.frx":1176
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   5
               Left            =   1620
               TabIndex        =   14
               Tag             =   "enabled"
               Top             =   2340
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":119E
               Caption         =   "frmManTipoCambio.frx":11BE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":122A
               Keys            =   "frmManTipoCambio.frx":1248
               Spin            =   "frmManTipoCambio.frx":12A8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   0
               Left            =   1620
               TabIndex        =   9
               Tag             =   "enabled"
               Top             =   270
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":12D0
               Caption         =   "frmManTipoCambio.frx":12F0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":135C
               Keys            =   "frmManTipoCambio.frx":137A
               Spin            =   "frmManTipoCambio.frx":13DA
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   1
               Left            =   1620
               TabIndex        =   10
               Tag             =   "enabled"
               Top             =   675
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1402
               Caption         =   "frmManTipoCambio.frx":1422
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":148E
               Keys            =   "frmManTipoCambio.frx":14AC
               Spin            =   "frmManTipoCambio.frx":150C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   2
               Left            =   1620
               TabIndex        =   11
               Tag             =   "enabled"
               Top             =   1080
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1534
               Caption         =   "frmManTipoCambio.frx":1554
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":15C0
               Keys            =   "frmManTipoCambio.frx":15DE
               Spin            =   "frmManTipoCambio.frx":163E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   9
               Left            =   5220
               TabIndex        =   18
               Tag             =   "enabled"
               Top             =   1485
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1666
               Caption         =   "frmManTipoCambio.frx":1686
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":16F2
               Keys            =   "frmManTipoCambio.frx":1710
               Spin            =   "frmManTipoCambio.frx":1770
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   10
               Left            =   5220
               TabIndex        =   19
               Tag             =   "enabled"
               Top             =   1890
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1798
               Caption         =   "frmManTipoCambio.frx":17B8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":1824
               Keys            =   "frmManTipoCambio.frx":1842
               Spin            =   "frmManTipoCambio.frx":18A2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   11
               Left            =   5220
               TabIndex        =   20
               Tag             =   "enabled"
               Top             =   2340
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":18CA
               Caption         =   "frmManTipoCambio.frx":18EA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":1956
               Keys            =   "frmManTipoCambio.frx":1974
               Spin            =   "frmManTipoCambio.frx":19D4
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   6
               Left            =   5220
               TabIndex        =   15
               Tag             =   "enabled"
               Top             =   270
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":19FC
               Caption         =   "frmManTipoCambio.frx":1A1C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":1A88
               Keys            =   "frmManTipoCambio.frx":1AA6
               Spin            =   "frmManTipoCambio.frx":1B06
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   7
               Left            =   5220
               TabIndex        =   16
               Tag             =   "enabled"
               Top             =   675
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1B2E
               Caption         =   "frmManTipoCambio.frx":1B4E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":1BBA
               Keys            =   "frmManTipoCambio.frx":1BD8
               Spin            =   "frmManTipoCambio.frx":1C38
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbcTCMes 
               Height          =   300
               Index           =   8
               Left            =   5220
               TabIndex        =   17
               Tag             =   "enabled"
               Top             =   1080
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":1C60
               Caption         =   "frmManTipoCambio.frx":1C80
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":1CEC
               Keys            =   "frmManTipoCambio.frx":1D0A
               Spin            =   "frmManTipoCambio.frx":1D6A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "SETIEMBRE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   17
               Left            =   3735
               TabIndex        =   56
               Top             =   1170
               Width           =   1140
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "OCTUBRE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   16
               Left            =   3735
               TabIndex        =   55
               Top             =   1530
               Width           =   960
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "AGOSTO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   15
               Left            =   3735
               TabIndex        =   54
               Top             =   720
               Width           =   870
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "DICIEMBRE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   14
               Left            =   3735
               TabIndex        =   53
               Top             =   2385
               Width           =   1140
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "JULIO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   13
               Left            =   3735
               TabIndex        =   52
               Top             =   330
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "NOVIEMBRE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   12
               Left            =   3735
               TabIndex        =   51
               Top             =   1935
               Width           =   1215
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "MARZO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   11
               Left            =   135
               TabIndex        =   50
               Top             =   1170
               Width           =   720
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "ABRIL"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   10
               Left            =   135
               TabIndex        =   49
               Top             =   1530
               Width           =   615
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "FEBRERO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   7
               Left            =   135
               TabIndex        =   48
               Top             =   720
               Width           =   930
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "JUNIO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   6
               Left            =   135
               TabIndex        =   46
               Top             =   2385
               Width           =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "ENERO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   4
               Left            =   135
               TabIndex        =   45
               Top             =   330
               Width           =   690
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "MAYO"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   44
               Top             =   1935
               Width           =   615
            End
         End
         Begin TrueOleDBList70.TDBCombo tdbcTipoMensual 
            Height          =   300
            Left            =   2655
            TabIndex        =   58
            Top             =   315
            Width           =   2895
            _ExtentX        =   5106
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
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=688"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=609"
            Splits(0)._ColumnProps(11)=   "Column(1)._VertColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
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
            _PropDict       =   $"frmManTipoCambio.frx":1D92
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
         Begin TrueOleDBList70.TDBCombo tdbcDe 
            Height          =   300
            Index           =   1
            Left            =   2655
            TabIndex        =   59
            Tag             =   "enabled"
            Top             =   765
            Width           =   2895
            _ExtentX        =   5106
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
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
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
            _PropDict       =   $"frmManTipoCambio.frx":1E19
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
         Begin MSForms.CommandButton cmdGrabar 
            Height          =   390
            Left            =   2790
            TabIndex        =   60
            Top             =   4140
            Width           =   1665
            Caption         =   " Grabar"
            PicturePosition =   327683
            Size            =   "2937;688"
            Picture         =   "frmManTipoCambio.frx":1EA0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   1485
            TabIndex        =   57
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   1485
            TabIndex        =   47
            Top             =   840
            Width           =   750
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4860
         Left            =   -74730
         TabIndex        =   37
         Top             =   540
         Width           =   7395
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
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
            Left            =   3645
            TabIndex        =   5
            Top             =   420
            Width           =   870
         End
         Begin TDBDate6Ctl.TDBDate dtpFechaBus 
            Height          =   300
            Index           =   1
            Left            =   4545
            TabIndex        =   6
            Top             =   405
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   529
            Calendar        =   "frmManTipoCambio.frx":243A
            Caption         =   "frmManTipoCambio.frx":253C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoCambio.frx":25A0
            Keys            =   "frmManTipoCambio.frx":25BE
            Spin            =   "frmManTipoCambio.frx":262A
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
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   405
            Width           =   2265
            _ExtentX        =   3995
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
            _PropDict       =   $"frmManTipoCambio.frx":2652
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
         Begin TrueOleDBGrid70.TDBGrid tdbgMoneda 
            Height          =   3345
            Index           =   1
            Left            =   105
            TabIndex        =   8
            Top             =   1395
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   5900
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fecha"
            Columns(0).DataField=   "Tca_dFecha"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "Tca_cCodigoDestino"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Moneda"
            Columns(2).DataField=   "MonedaLargo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Venta Publicación"
            Columns(3).DataField=   "Tca_nVentaP"
            Columns(3).NumberFormat=   "External Editor"
            Columns(3).ExternalEditor=   "TDBNumber1"
            Columns(3).ExternalEditor.vt=   8
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Compra Vigente"
            Columns(4).DataField=   "Tca_nCompra"
            Columns(4).NumberFormat=   "External Editor"
            Columns(4).ExternalEditor=   "TDBNumber1"
            Columns(4).ExternalEditor.vt=   8
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Venta Vigente"
            Columns(5).DataField=   "Tca_nVenta"
            Columns(5).NumberFormat=   "External Editor"
            Columns(5).ExternalEditor=   "TDBNumber1"
            Columns(5).ExternalEditor.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=450"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=370"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=3413"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3334"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2566"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2487"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=530"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2090"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2011"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=530"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2143"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2064"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=530"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            HeadLines       =   2
            FootLines       =   1
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   0
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
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34,.bgcolor=&H8000000F&"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList70.TDBCombo tdbcMonedaAdic 
            Height          =   300
            Left            =   945
            TabIndex        =   7
            Top             =   855
            Width           =   5145
            _ExtentX        =   9075
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
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=688"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=609"
            Splits(0)._ColumnProps(11)=   "Column(1)._VertColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
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
            _PropDict       =   $"frmManTipoCambio.frx":26D9
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
         Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
            Height          =   300
            Left            =   5520
            TabIndex        =   39
            Tag             =   "enabled"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   529
            Calculator      =   "frmManTipoCambio.frx":2760
            Caption         =   "frmManTipoCambio.frx":2780
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoCambio.frx":27EC
            Keys            =   "frmManTipoCambio.frx":280A
            Spin            =   "frmManTipoCambio.frx":2862
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0.000"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "####0.000"
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1380909061
            MinValueVT      =   1162608645
         End
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
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
            Left            =   90
            TabIndex        =   41
            Top             =   900
            Width           =   660
         End
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
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
            Left            =   90
            TabIndex        =   40
            Top             =   450
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4455
         Index           =   0
         Left            =   -74775
         TabIndex        =   29
         Top             =   540
         Width           =   7095
         Begin VB.Frame Frame3 
            Caption         =   "Tipos de Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1890
            Index           =   0
            Left            =   1440
            TabIndex        =   30
            Top             =   2205
            Width           =   4215
            Begin TDBNumber6Ctl.TDBNumber tdbnCompra 
               Height          =   300
               Left            =   2265
               TabIndex        =   23
               Tag             =   "enabled"
               Top             =   330
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":288A
               Caption         =   "frmManTipoCambio.frx":28AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":2916
               Keys            =   "frmManTipoCambio.frx":2934
               Spin            =   "frmManTipoCambio.frx":2994
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnVenta 
               Height          =   300
               Left            =   2265
               TabIndex        =   24
               Tag             =   "enabled"
               Top             =   810
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":29BC
               Caption         =   "frmManTipoCambio.frx":29DC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":2A48
               Keys            =   "frmManTipoCambio.frx":2A66
               Spin            =   "frmManTipoCambio.frx":2AC6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnVentaP 
               Height          =   300
               Left            =   2280
               TabIndex        =   25
               Tag             =   "enabled"
               Top             =   1320
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   529
               Calculator      =   "frmManTipoCambio.frx":2AEE
               Caption         =   "frmManTipoCambio.frx":2B0E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmManTipoCambio.frx":2B7A
               Keys            =   "frmManTipoCambio.frx":2B98
               Spin            =   "frmManTipoCambio.frx":2BF8
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0.000"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0.000"
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1380909061
               MinValueVT      =   1162608645
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Venta Vigente"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   3
               Left            =   105
               TabIndex        =   33
               Top             =   810
               Width           =   1380
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Compra Vigente"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   2
               Left            =   105
               TabIndex        =   32
               Top             =   330
               Width           =   1575
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Venta Publicación"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   210
               Index           =   5
               Left            =   120
               TabIndex        =   31
               Top             =   1320
               Width           =   1755
            End
         End
         Begin TDBDate6Ctl.TDBDate dtpFecha 
            Height          =   300
            Left            =   2685
            TabIndex        =   22
            Tag             =   "_"
            Top             =   1410
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Calendar        =   "frmManTipoCambio.frx":2C20
            Caption         =   "frmManTipoCambio.frx":2D22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoCambio.frx":2D86
            Keys            =   "frmManTipoCambio.frx":2DA4
            Spin            =   "frmManTipoCambio.frx":2E10
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
         Begin TrueOleDBList70.TDBCombo tdbcDe 
            Height          =   300
            Index           =   0
            Left            =   2685
            TabIndex        =   21
            Tag             =   "_"
            Top             =   900
            Width           =   2895
            _ExtentX        =   5106
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
            _PropDict       =   $"frmManTipoCambio.frx":2E38
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
            Left            =   2745
            TabIndex        =   36
            Top             =   225
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   1530
            TabIndex        =   35
            Top             =   1395
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1530
            TabIndex        =   34
            Top             =   930
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4860
         Left            =   270
         TabIndex        =   27
         Top             =   540
         Width           =   7395
         Begin VB.CheckBox chkFecha 
            Caption         =   "Fecha"
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
            Left            =   4050
            TabIndex        =   1
            Top             =   420
            Width           =   870
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgMoneda 
            Height          =   3840
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   900
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   6773
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fecha"
            Columns(0).DataField=   "Tca_dFecha"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "Tca_cCodigoDestino"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Moneda"
            Columns(2).DataField=   "MonedaLargo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Venta Publicación"
            Columns(3).DataField=   "Tca_nVentaP"
            Columns(3).NumberFormat=   "External Editor"
            Columns(3).ExternalEditor=   "TDBNumber1"
            Columns(3).ExternalEditor.vt=   8
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Compra Vigente"
            Columns(4).DataField=   "Tca_nCompra"
            Columns(4).NumberFormat=   "External Editor"
            Columns(4).ExternalEditor=   "TDBNumber1"
            Columns(4).ExternalEditor.vt=   8
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Venta Vigente"
            Columns(5).DataField=   "Tca_nVenta"
            Columns(5).NumberFormat=   "External Editor"
            Columns(5).ExternalEditor=   "TDBNumber1"
            Columns(5).ExternalEditor.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=450"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=370"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2381"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2566"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2487"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=530"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2090"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2011"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=530"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2143"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2064"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=530"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            HeadLines       =   2
            FootLines       =   1
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   0
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
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34,.bgcolor=&H8000000F&"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   300
            Left            =   5520
            TabIndex        =   28
            Tag             =   "enabled"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   529
            Calculator      =   "frmManTipoCambio.frx":2EBF
            Caption         =   "frmManTipoCambio.frx":2EDF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoCambio.frx":2F4B
            Keys            =   "frmManTipoCambio.frx":2F69
            Spin            =   "frmManTipoCambio.frx":2FC1
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0.000"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "####0.000"
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1380909061
            MinValueVT      =   1162608645
         End
         Begin TDBDate6Ctl.TDBDate dtpFechaBus 
            Height          =   300
            Index           =   0
            Left            =   5040
            TabIndex        =   2
            Top             =   405
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   529
            Calendar        =   "frmManTipoCambio.frx":2FE9
            Caption         =   "frmManTipoCambio.frx":30EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoCambio.frx":314F
            Keys            =   "frmManTipoCambio.frx":316D
            Spin            =   "frmManTipoCambio.frx":31D9
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
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   0
            Top             =   405
            Width           =   2850
            _ExtentX        =   5027
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
            _PropDict       =   $"frmManTipoCambio.frx":3201
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
         Begin VB.Label lblPeriodo 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
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
            Left            =   90
            TabIndex        =   38
            Top             =   450
            Width           =   645
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmManTipoCambio.frx":3288
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":3662
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":3A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":3E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":41F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":45CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":49A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":4D7E
            Key             =   ""
         EndProperty
      EndProperty
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
            Picture         =   "frmManTipoCambio.frx":5D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":5EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":604C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":61A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":6300
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":645A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":65B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":670E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":6868
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
            Picture         =   "frmManTipoCambio.frx":69C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":6F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":74F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":7A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":802A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":85C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":8B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":90F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoCambio.frx":9692
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   61
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
Attribute VB_Name = "frmManTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla(1) As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim gsPeriodoAnterior(1) As String
Dim IndiceTab As Integer
Dim IndiceTabAnt As Integer
Public Voucher As Boolean
Public RegAux  As Boolean
Public MonAdic As Boolean
Public ColumnaTC As Integer
Public Asientos As Boolean


Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkFecha_Click(Index As Integer)
    Call FiltrarRecordSet(Index)
End Sub

Private Sub chkFecha_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    'pSetFocus dtpFechaBus
End If

End Sub



Private Sub dtpFechaBus_Change(Index As Integer)
    If chkFecha(Index).Value = 1 Then Call FiltrarRecordSet(Index)
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left + 15 - 200
            .Height = Me.Height - .Top + 15 - 500
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 500
        End With
       
        With tdbgMoneda(0)
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        With tdbgMoneda(1)
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        '*** REDIMENSIONAR DETALLE
        Frame4.Height = Frame1.Height
        Frame4.Width = Frame1.Width
        
        Frame2(0).Height = Frame1.Height
        Frame2(0).Width = Frame1.Width
        
        Frame2(1).Height = Frame1.Height
        Frame2(1).Width = Frame1.Width
        
        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
End Sub

Private Sub SSTCentroCosto_Click(PreviousTab As Integer)
    If PreviousTab < 2 Then IndiceTabAnt = PreviousTab
    IndiceTab = SSTCentroCosto.Tab
    
    If SSTCentroCosto.Tab = 2 Or SSTCentroCosto.Tab = 4 Then
        tbrOpciones.Buttons(1).Enabled = False
        tbrOpciones.Buttons(3).Enabled = False
        tbrOpciones.Buttons(4).Enabled = False
        tbrOpciones.Buttons(5).Enabled = False
    
    ElseIf SSTCentroCosto.Tab = 0 Or SSTCentroCosto.Tab = 1 Then
        tbrOpciones.Buttons(1).Enabled = True
        tbrOpciones.Buttons(3).Enabled = True
        tbrOpciones.Buttons(4).Enabled = True
        tbrOpciones.Buttons(5).Enabled = True
        
        SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    End If
    
    If SSTCentroCosto.Tab = 2 Then
        Call CargaMeses
    End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    If IndiceTab < 2 Then IndiceTabAnt = IndiceTab
    
    IndiceTab = SSTCentroCosto.Tab
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo
        Case 2: VerDatos
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                
        Case 4: Borrar (IndiceTab)
        Case 5: Editar
        Case 6: Imprimir
        Case 7
            If SSTCentroCosto.TabEnabled(3) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
            End If
    End Select
End Sub

Private Sub Borrar(Indice As Integer)
    If SSTCentroCosto.Tab = 2 Then
        Exit Sub
    End If

    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgMoneda(Indice).Columns(0).Value) <> "" Then
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            Call CargaArregloMnt
            lArrMnt(0) = "ELIMINAR"                     ' Accion
            lArrMnt(2) = tdbgMoneda(Indice).Columns(0).Value  ' Fecha
            lArrMnt(3) = gsMonedaNac     ' MonedaOrigen
            lArrMnt(4) = tdbgMoneda(Indice).Columns(1).Value
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoCambio", lArrMnt(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Call CargaTablaMoneda(IndiceTabAnt)
            Screen.MousePointer = vbDefault
            FiltrarRecordSet (IndiceTabAnt)
            Mensajes "Registro ha sido eliminado", vbInformation
        End If
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

Private Sub VerDatos()
    lblMante = "VER REGISTRO"
    Call CargaDatosRegistro(IndiceTab)
    SSTCentroCosto.TabEnabled(3) = True
    SSTCentroCosto.TabEnabled(0) = False
    SSTCentroCosto.Tab = 3
    tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
    tbrOpciones.Buttons(7).Image = 8
    lTipoMnt = "EDITAR"
    Call AseguraControl(Me, True)
End Sub

Private Sub Editar()
    If SSTCentroCosto.Tab = 2 Then
        Exit Sub
    End If

    Call CargaDatosRegistro(IndiceTab)
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
        dtpFecha.ReadOnly = True
        'tdbcA.Locked = True
        tdbcDe(0).Locked = True
        pSetFocus tdbnCompra
       
    Else
        lRegElim = False
    End If
End Sub

Public Sub ManNuevo()
    On Error GoTo ERROR
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    dtpFecha.ReadOnly = False
    'tdbcA.Locked = False
    tdbcDe(0).Locked = False
    dtpFecha = FechaServidor
    pSetFocus dtpFecha
    
    If IndiceTabAnt = 0 Then
        tdbcDe(0).BoundText = gsMonedaExt
        tdbcDe(0).Locked = True
    Else
        tdbcDe(0).BoundText = tdbcMonedaAdic.BoundText
        tdbcDe(0).Locked = False
    End If
    If Me.tdbgMoneda(IndiceTabAnt).Row > 0 Then Me.tdbgMoneda(IndiceTabAnt).Row = 0
    If CE(Me.tdbgMoneda(IndiceTabAnt).Columns(0)) <> "" Then
        dtpFecha.Value = DateAdd("d", 1, Me.tdbgMoneda(IndiceTabAnt).Columns(0))
        Me.tdbcMes(IndiceTabAnt).BoundText = Right("00" & Month(dtpFecha), 2)
    Else
        dtpFecha.Value = "01/" + tdbcMes(IndiceTabAnt).BoundText + "/" + gsAnio
        
    End If
    
    tdbnCompra.Enabled = True
    tdbnVenta.Enabled = True
    tdbnVentaP.Enabled = True
    pSetFocus tdbnCompra
    Exit Sub
ERROR:
    
End Sub

Private Sub BuscarMonedas(Indice As Integer)
    On Error GoTo ERROR
    ' *** Busca la moneda nacional y extranjera por defecto
    Dim i As Integer
    Dim Cont As Integer
    Cont = 0
    For i = 0 To tdbcDe(Indice).ListCount - 1
        tdbcDe(Indice).Row = i
        If tdbcDe(Indice).Columns(2) = "1" Then
            tdbcDe(Indice).Bookmark = i
            Cont = Cont + 1
        End If
        If tdbcDe(Indice).Columns(3) = "1" Then
            'tdbcA.Bookmark = i
            Cont = Cont + 1
        End If
        If Cont = 2 Then Exit For
    Next
    Exit Sub
ERROR:
    
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    SSTCentroCosto.TabEnabled(3) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    SSTCentroCosto.TabEnabled(1) = Not Valor
    SSTCentroCosto.TabEnabled(2) = Not Valor
    'SSTCentroCosto.TabEnabled(4) = Not Valor
    
    If Valor = True Then SSTCentroCosto.Tab = 3
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
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
    SSTCentroCosto.Tab = IndiceTabAnt
    pSetFocus tdbgMoneda(IndiceTabAnt)
End Sub

Private Function BuscaFecha() As Date
    Select Case SSTCentroCosto.Tab
           Case 0
                BuscaFecha = dtpFechaBus(0).Value
           Case 1
                BuscaFecha = dtpFechaBus(0).Value
    End Select
End Function

Private Function BuscaMoneda() As String
    Select Case SSTCentroCosto.Tab
           Case 0
                BuscaMoneda = gsMonedaExt
           Case 1
                BuscaMoneda = tdbcMonedaAdic.BoundText
    End Select
End Function

Private Sub Imprimir()
    Dim matriz(15) As Variant
    Dim Titulo As String
    
    
    If SSTCentroCosto.Tab = 0 Then
        Titulo = "Tipo de Cambio - " & gsNombreMonedaExt
        
    ElseIf SSTCentroCosto.Tab = 1 Then
        If tdbcMonedaAdic.BoundText = "" Then
            Mensajes "Seleccione una moneda Adicional", vbOKOnly + vbInformation
            Exit Sub
        End If
        Titulo = "Tipo de Cambio - " & tdbcMonedaAdic.Text
        
    ElseIf SSTCentroCosto.Tab = 2 Then
        Titulo = "Tipo de Cambio Mensual " & gsNombreMonedaExt & " " & gsAnio
    End If
    
    
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;;True"
    matriz(3) = "@Titulo02;;True"
    matriz(4) = "@Titulo03;FECHA;True"
    matriz(5) = "@Titulo04;COMPRA VIG.;True"
    matriz(6) = "@Titulo05;VENTA VIG.;True"
    matriz(7) = "@Titulo06;VENTA PUBL.;True"
    matriz(8) = "@Titulo07;;True"
    
    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    If SSTCentroCosto.Tab <= 1 Then
        matriz(9) = "@Tipo;TIPO_CAMBIO;True"
    Else
        matriz(9) = "@Tipo;TIPO_CAMBIO_MENSUAL;True"
    End If
    
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    
    Dim formulas(0) As Variant
    
    If SSTCentroCosto.Tab = 0 Then
        matriz(12) = "@Per_cPeriodo;" & tdbcMes(0).BoundText & ";True"
        matriz(13) = "@Aux;" & gsMonedaExt & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandar.rpt", crptToWindow, Titulo, "", matriz(), formulas()
        
    ElseIf SSTCentroCosto.Tab = 1 Then
        matriz(12) = "@Per_cPeriodo;" & tdbcMes(1).BoundText & ";True"
        matriz(13) = "@Aux;" & tdbcMonedaAdic.BoundText & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandar.rpt", crptToWindow, Titulo, "", matriz(), formulas()

    ElseIf SSTCentroCosto.Tab = 2 Then
        matriz(12) = "@Per_cPeriodo;;True"
        matriz(13) = "@Aux;" & tdbcTipoMensual.BoundText & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptTipoCambioMensual.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
    End If
    
    
    

End Sub

Private Sub Grabar()

    If CE(tdbcDe(0).BoundText) = "" Then
        Mensajes "Seleccione un tipo de moneda la lista", vbOKOnly + vbInformation
        Exit Sub
    End If


    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    If validarDatos = False Then Exit Sub
    Set clsMante = New clsMantoTablas

    On Local Error GoTo ErrorEjecucion
    
    Call CargaArregloMnt
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoCambio", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    '-----------------------------------------------------------'
    
    Dim cMes As String
    cMes = Right("00" & dtpFecha.Month, 2)
    
    If UltimoDiaMes(cMes, gsAnio) = dtpFecha.Value Then
        Call GrabaTcMensual(cMes, tdbnCompra.Value, TCM_COMPRA)
        Call GrabaTcMensual(cMes, tdbnVenta.Value, TCM_VENTA)
    End If
    '-----------------------------------------------------------'
    Call Cancelar
    CargaTablaMoneda (IndiceTab)
    FiltrarRecordSet (IndiceTab)
    On Error Resume Next
    lrsTabla(IndiceTab).Find "Tca_dFecha = '" & dtpFecha & "'"

    Mensajes "Los datos se grabaron con exito...", vbInformation + vbOKOnly
    tdbgMoneda(IndiceTab).HighlightRowStyle = "HighlightRow"
    
    If Voucher = True Then 'si fue llamado del mant de voucher asignarle el TC VEP a la celda
        On Error Resume Next
        frmManAsientosContables.Enabled = True
        frmManAsientosContables.tdbgDetalle.Columns(16) = Me.tdbnVentaP
        Unload Me
        pSetFocus frmManAsientosContables.tdbgDetalle
        pSendKeys "{Enter}"
        On Error GoTo 0
    End If
    
    If MonAdic = True Then  'si fue llamado del mant de voucher asignarle el TC VEP a la celda
        On Error Resume Next
        frmManAsientosContables.Enabled = True
        frmManAsientosContables.tdbgDetalle.Columns(ColumnaTC) = Me.tdbnVentaP.Value * frmManAsientosContables.tdbtMonedaAdic.Value
        frmManAsientosContables.ValoresMonedaAdic
        Unload Me
        pSetFocus frmManAsientosContables.tdbgDetalle
        

        pSendKeys "{Enter}"
        On Error GoTo 0
    End If
    
    If RegAux = True Then  'si fue llamado de registro de auxiliares
        On Error Resume Next
        FrmManRegAuxiliarVentas.Enabled = True
        FrmManRegAuxiliarVentas.tdbTC.Value = Me.tdbnVentaP.Value
        Unload Me
        pSetFocus FrmManRegAuxiliarVentas.tdbTC

        pSendKeys "{Enter}"
        On Error GoTo 0
    End If
    Exit Sub
    
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If lTipoMnt = "INSERTAR" Then
        If ExisteCambio = True Then Exit Function
    End If
   
    ' ***
    validarDatos = True
End Function

Private Function ExisteCambio() As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    ExisteCambio = False
    sqlSp = "spCn_GrabaTipoCambio 'SEL_REG', '" & gsEmpresa & "', '" & dtpFecha.Value & "', '" & gsMonedaNac & "', '" & tdbcDe(0).BoundText & "', 0, 0,0, 0, '' "
        arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteCambio = True
        Mensajes "Cambio con fecha indicada ya existe. Verifique...", vbInformation
        pSetFocus dtpFecha
    End If
    Call CerrarRecordSet(rsArreglo)
End Function

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(10) As Variant
    lArrMnt(0) = lTipoMnt           ' Accion
    lArrMnt(1) = gsEmpresa          ' Empresa
    lArrMnt(2) = dtpFecha           ' Fecha
    lArrMnt(3) = gsMonedaNac        ' MonedaOrigen
    lArrMnt(4) = tdbcDe(0).BoundText ' MonedaDestino
    lArrMnt(5) = tdbnCompra.Value          ' Compra
    lArrMnt(6) = tdbnVenta.Value           ' Venta
    lArrMnt(7) = 0                  ' Compra Publicación
    lArrMnt(8) = tdbnVentaP         ' Venta Publicación
    lArrMnt(9) = gsUsuario          ' Usuario
    lArrMnt(10) = gsPeriodo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(3) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar (IndiceTab)
        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub

Private Sub Form_Load()
   On Error GoTo ERROR
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    Dim Mes As String
    Asientos = False
    Mes = gsPeriodo
   
    If Mes = "00" Then Mes = "01"
    If Mes > "12" Then Mes = "12"
    
    
    Me.MonAdic = False 'variable que indica si fue llamado del mant de voucher (mon adicional)
    Me.Voucher = False 'variable que indica si fue llamado del mant de voucher
    
    Call Centrar_form(Me)
    
    
    dtpFecha.MinDate = "01/01/1000"
    dtpFecha.MaxDate = "31/12/2500"

    dtpFechaBus(0).MinDate = "01/01/1000"
    dtpFechaBus(0).MaxDate = "31/12/2500"

    dtpFechaBus(1).MinDate = "01/01/1000"
    dtpFechaBus(1).MaxDate = "31/12/2500"

    
    Call LlenaCombos
    Call LlenaComboMesActivo(tdbcMes(0))
    Call LlenaComboMesActivo(tdbcMes(1))
    
    tdbcDe(1).Locked = True
    
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(3) = False
    
    lRegElim = False
    
    
    IndiceTab = 0
    tdbgMoneda(0).HighlightRowStyle = "HighlightRow"
    tdbgMoneda(1).HighlightRowStyle = "HighlightRow"
    
    Call CargaMeses
'    Call BuscarTCDAOT
    On Error Resume Next
    
    tdbcTipoMensual.BoundText = "1"
    tdbcDe(1).BoundText = gsMonedaExt
    
    
    tdbcMes(0).BoundText = Mes
    tdbcMes(1).BoundText = Mes
   
    tdbcMes(0).ReBind
    tdbcMes(1).ReBind
    
    tdbcDe(0).ReBind
    tdbcDe(1).ReBind

    tdbcTipoMensual.ReBind
    tdbcMonedaAdic.ReBind
   
    DoEvents
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    SSTCentroCosto.Tab = 0
    
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        cmdGrabar.Enabled = False
    Else
        cmdGrabar.Enabled = True
    End If
    
    
   Exit Sub
    
ERROR:
    Mensajes Err.Description, vbOKOnly + vbCritical
    
End Sub

Private Sub SeteaBarraHerram()

End Sub

Private Sub BuscarTCDAOT()

End Sub



Public Sub ConfigurarControlFecha(Indice As Integer)
   On Error GoTo ERROR
   Dim FechaIni As Date, FechaFin As Date, NuevaFecha As Date
   Dim Mes As String
   Mes = gsPeriodo
   If Mes < "01" Then Mes = "01"
   If Mes > "12" Then Mes = "12"
   
   
   FechaIni = dtpFechaBus(Indice).MinDate
   FechaFin = dtpFechaBus(Indice).MaxDate
   NuevaFecha = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
   DoEvents
   If Indice > 1 Then Exit Sub
    
    If Val(tdbcMes(Indice).BoundText) > 0 And Val(tdbcMes(Indice).BoundText) < 13 Then
        dtpFechaBus(Indice).Enabled = True
        
        
        If Format(NuevaFecha, "yyyyMMdd") <= Format(FechaIni, "yyyyMMdd") Then
            dtpFechaBus(Indice).MinDate = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
            dtpFechaBus(Indice).MaxDate = UltimoDiaMes(tdbcMes(Indice).BoundText, gsAnio)
        End If
        
        If Format(NuevaFecha, "yyyyMMdd") >= Format(FechaFin, "yyyyMMdd") Then
            dtpFechaBus(Indice).MaxDate = UltimoDiaMes(tdbcMes(Indice).BoundText, gsAnio)
            dtpFechaBus(Indice).MinDate = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
        End If
        
        If Val(tdbcMes(Indice).BoundText) = Val(Month(Date)) And gsAnio = Val(Year(Date)) Then
            dtpFechaBus(Indice).Value = Date
        Else
            dtpFechaBus(Indice).Value = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
        End If
        
        gsPeriodoAnterior(Indice) = tdbcMes(Indice).BoundText
        
        DoEvents
    Else
        dtpFechaBus(Indice).Enabled = False
        dtpFechaBus(Indice) = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
    End If
    DoEvents
    Exit Sub
ERROR:
    Mensajes Err.Description & Chr(10) + Chr(13) & _
            "Rango: " & dtpFechaBus(Indice).Value & Chr(10) + Chr(13) & _
            "Min  : " & dtpFechaBus(Indice).MinDate & Chr(10) + Chr(13) & _
            "Max  : " & dtpFechaBus(Indice).MaxDate, vbOKOnly + vbCritical
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla(0))
    Call CerrarRecordSet(lrsTabla(1))
    
    If Me.Voucher = True Or Me.MonAdic = True Then
        frmManAsientosContables.Enabled = True
        pSetFocus frmManAsientosContables.tdbgDetalle
    End If
    
    If RegAux = True Then
        FrmManRegAuxiliarVentas.Enabled = True
    End If
    
    If Asientos = True Then
        frmBusTipoAsiento.Enabled = True
        'frmBusTipoAsiento.Insertar
        frmBusTipoAsiento.CargaTC
    End If
    
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

Private Sub CargaMeses()
    Dim sql As String
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    sql = "select * from CNT_TIPO_CAMBIO_MENSUAL " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    If Not rsAddItem Is Nothing Then
        Do While Not rsAddItem.EOF
            tdbcTCMes(0) = NE(rsAddItem!Tca_cEne)
            tdbcTCMes(1) = NE(rsAddItem!Tca_cFeb)
            tdbcTCMes(2) = NE(rsAddItem!Tca_cMar)
            tdbcTCMes(3) = NE(rsAddItem!Tca_cAbr)
            tdbcTCMes(4) = NE(rsAddItem!Tca_cMay)
            tdbcTCMes(5) = NE(rsAddItem!Tca_cJun)
            tdbcTCMes(6) = NE(rsAddItem!Tca_cJul)
            tdbcTCMes(7) = NE(rsAddItem!Tca_cAgo)
            tdbcTCMes(8) = NE(rsAddItem!Tca_cSet)
            tdbcTCMes(9) = NE(rsAddItem!Tca_cOct)
            tdbcTCMes(10) = NE(rsAddItem!Tca_cNov)
            tdbcTCMes(11) = NE(rsAddItem!Tca_cDic)
            
            rsAddItem.MoveNext
        Loop
    Else
        tdbcTCMes(0) = 0
        tdbcTCMes(1) = 0
        tdbcTCMes(2) = 0
        tdbcTCMes(3) = 0
        tdbcTCMes(4) = 0
        tdbcTCMes(5) = 0
        tdbcTCMes(6) = 0
        tdbcTCMes(7) = 0
        tdbcTCMes(8) = 0
        tdbcTCMes(9) = 0
        tdbcTCMes(10) = 0
        tdbcTCMes(11) = 0
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
End Sub

Private Function Valida() As Boolean
    If CE(tdbcTipoMensual.BoundText) = "" Then
        Mensajes "Seleccione un tipo de moneda", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If CE(tdbcDe(1).BoundText) = "" Then
        Mensajes "Seleccione una moneda", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    Valida = True
End Function

Private Sub GrabaTcMensual(cMes As String, nvalor As Double, cTipo As Tipo_Cambio)

    Dim sql As String
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As New ClsFuncionesExecute
    Dim Existe As Boolean
    On Error GoTo ERROR
    Dim cCadena As String
    
    Select Case cMes
        Case "01": cCadena = "Tca_cEne"
        Case "02": cCadena = "Tca_cFeb"
        Case "03": cCadena = "Tca_cMar"
        Case "04": cCadena = "Tca_cAbr"
        Case "05": cCadena = "Tca_cMay"
        Case "06": cCadena = "Tca_cJun"
        Case "07": cCadena = "Tca_cJul"
        Case "08": cCadena = "Tca_cAgo"
        Case "09": cCadena = "Tca_cSet"
        Case "10": cCadena = "Tca_cOct"
        Case "11": cCadena = "Tca_cNov"
        Case "12": cCadena = "Tca_cDic"
    End Select
    
    sql = "select count(emp_ccodigo) as Registro from CNT_TIPO_CAMBIO_MENSUAL " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & cTipo & "' "
    
    Existe = False
    Set rsAddItem = clDatos.fRetornaRS(sql)
    If Not rsAddItem Is Nothing Then
        Do While Not rsAddItem.EOF
            If NE(rsAddItem!Registro) > 0 Then
                Existe = True
            End If
            rsAddItem.MoveNext
        Loop
    End If
    
    If Existe = True Then
        sql = "Update CNT_TIPO_CAMBIO_MENSUAL set " & cCadena & "=" & nvalor & " " & _
              "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
              "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & cTipo & "'"

    Else
        sql = "Insert into CNT_TIPO_CAMBIO_MENSUAL (" & cCadena & "," & _
                                                    "emp_ccodigo,pan_canio,tca_cmoneda,tca_ctipo) values (" & _
                                                    nvalor & "," & _
                                                    "'" & gsEmpresa & "','" & gsAnio & "','" & gsMonedaExt & "'," & _
                                                    "'" & cTipo & "')"
    End If

    clDatos.pEjecutaSQL (sql)
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    
    Exit Sub
    
ERROR:
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing

End Sub
Private Sub cmdGrabar_Click()
    If Valida = False Then Exit Sub

    Dim sql As String
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As New ClsFuncionesExecute
    Dim Existe As Boolean
    Screen.MousePointer = vbHourglass
    On Error GoTo ERROR
    sql = "select count(*) as Registro from CNT_TIPO_CAMBIO_MENSUAL " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"
    
    Existe = False
    Set rsAddItem = clDatos.fRetornaRS(sql)
    If Not rsAddItem Is Nothing Then
        Do While Not rsAddItem.EOF
            If NE(rsAddItem!Registro) > 0 Then
                Existe = True
            End If
            rsAddItem.MoveNext
        Loop
    End If
    
    If Existe = True Then
        sql = "Update CNT_TIPO_CAMBIO_MENSUAL set Tca_cEne=" & NE(tdbcTCMes(0)) & "," & _
                                                 "Tca_cFeb=" & NE(tdbcTCMes(1)) & "," & _
                                                 "Tca_cMar=" & NE(tdbcTCMes(2)) & "," & _
                                                 "Tca_cAbr=" & NE(tdbcTCMes(3)) & "," & _
                                                 "Tca_cMay=" & NE(tdbcTCMes(4)) & "," & _
                                                 "Tca_cJun=" & NE(tdbcTCMes(5)) & "," & _
                                                 "Tca_cJul=" & NE(tdbcTCMes(6)) & "," & _
                                                 "Tca_cAgo=" & NE(tdbcTCMes(7)) & "," & _
                                                 "Tca_cSet=" & NE(tdbcTCMes(8)) & "," & _
                                                 "Tca_cOct=" & NE(tdbcTCMes(9)) & "," & _
                                                 "Tca_cNov=" & NE(tdbcTCMes(10)) & "," & _
                                                 "Tca_cDic=" & NE(tdbcTCMes(11)) & " " & _
              "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
              "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"

    Else
        sql = "Insert into CNT_TIPO_CAMBIO_MENSUAL (Tca_cEne,Tca_cFeb,Tca_cMar,Tca_cAbr," & _
                                                    "Tca_cMay,Tca_cJun,Tca_cJul,Tca_cAgo," & _
                                                    "Tca_cSet,Tca_cOct,Tca_cNov,Tca_cDic," & _
                                                    "emp_ccodigo,pan_canio,tca_cmoneda,tca_ctipo) values (" & _
                                                     NE(tdbcTCMes(0)) & "," & NE(tdbcTCMes(1)) & "," & _
                                                     NE(tdbcTCMes(2)) & "," & NE(tdbcTCMes(3)) & "," & _
                                                     NE(tdbcTCMes(4)) & "," & NE(tdbcTCMes(5)) & "," & _
                                                     NE(tdbcTCMes(6)) & "," & NE(tdbcTCMes(7)) & "," & _
                                                     NE(tdbcTCMes(8)) & "," & NE(tdbcTCMes(9)) & "," & _
                                                     NE(tdbcTCMes(10)) & "," & NE(tdbcTCMes(11)) & "," & _
                                                    "'" & gsEmpresa & "','" & gsAnio & "','" & gsMonedaExt & "'," & _
                                                    "'" & Me.tdbcTipoMensual.BoundText & "')"
    End If

    clDatos.pEjecutaSQL (sql)
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    
    CargaMeses
    Screen.MousePointer = vbNormal
    Mensajes "Se grabo correctamente los tipos de cambios mensuales", vbOKOnly + vbInformation
    
    Exit Sub
    
ERROR:
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    Screen.MousePointer = vbNormal
    Mensajes "No se grabo correctamente los tipos de cambios mensuales", vbOKOnly + vbInformation

End Sub

Private Sub LlenaCombos()
    On Error GoTo ERROR
    Dim sqlcombos As String

    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMNac<> '1' " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcDe(0), sqlcombos

    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMExt = '1' " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcDe(1), sqlcombos

    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt, Mon_cNombreCorto From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMNac <> '1' and Mon_cMExt <> '1' " & _
                "ORDER BY Mon_cNombreLargo"
                
    LlenarComboAddItem tdbcMonedaAdic, sqlcombos, True

    sqlcombos = "select Tab_cCodigo, Tab_cDescripCampo from tabla  " & _
                "where emp_ccodigo='" & gsEmpresa & "' and tab_ctabla='046' AND Tab_cDescripCampo LIKE 'CIERRE%' " & _
                "ORDER BY Tab_cCodigo"
                
    LlenarComboAddItem tdbcTipoMensual, sqlcombos
    Exit Sub
ERROR:

End Sub

Private Sub CargaTablaMoneda(Indice As Integer)
    If Indice > 1 Then Exit Sub
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
    Set lrsTabla(Indice) = New ADODB.Recordset
    Set tdbgMoneda(Indice).DataSource = Nothing
    Dim fechaAux As String
    Dim Moneda As String
    
    If Indice = 0 Then
        Moneda = gsMonedaExt
    Else
        Moneda = tdbcMonedaAdic.BoundText
    End If
    
    fechaAux = "01/" & Right("00" & dtpFechaBus(Indice).Month, 2) & "/" & gsAnio
    sqlSp = "spCn_GrabaTipoCambio 'SEL_ALL', '" & gsEmpresa & "', '" & fechaAux & "', '" & gsMonedaNac & "', '" & Moneda & "', 0, 0,0, 0, ''"
    arrDatos = Array(sqlSp)
    Set lrsTabla(Indice) = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla(Indice) Is Nothing Then
        If Not (lrsTabla(Indice).EOF And lrsTabla(Indice).BOF) Then

        lrsTabla(Indice).Sort = "Tca_dFecha desc"
        tdbgMoneda(Indice).DataSource = lrsTabla(Indice)

        End If
    End If
End Sub

Private Sub CargaDatosRegistro(Indice As Integer)
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    Dim sMoneda As String
    
    With tdbgMoneda(Indice)
        On Error GoTo serror
        
        sMoneda = .Columns(1).Value
        
        If SSTCentroCosto.Tab = 0 Then sMoneda = gsMonedaExt
        If SSTCentroCosto.Tab = 1 Then sMoneda = tdbcMonedaAdic.BoundText
        
        sqlSp = "spCn_GrabaTipoCambio 'SEL_REG', '" & gsEmpresa & "', '" & .Columns(0).Value & "', '" & gsMonedaNac & "','" & sMoneda & "', 0, 0,0, 0, '' "
    End With
    
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        'Mensajes "Seleccione un tipo de moneda", vbInformation
        Set rsArreglo = Nothing
        'Exit Sub
    End If
    ' *** Asignando Datos del Tipo de Cambio
    dtpFecha = CE(rsArreglo!Tca_dFecha)
    
    tdbcDe(0).BoundText = CE(rsArreglo!Tca_cCodigoDestino)
    
    'tdbcA.BoundText = rsArreglo!Tca_cCodigoDestino
    tdbnCompra = NE(rsArreglo!Tca_nCompra)
    tdbnVenta = NE(rsArreglo!Tca_nVenta)
    'tdbnCompraP = NuloNum(rsArreglo!Tca_nCompraP)  Operacion
    tdbnVentaP = NE(rsArreglo!Tca_nVentaP)
    Call CerrarRecordSet(rsArreglo)
    ' ***
    Exit Sub
serror:
    Mensajes "Seleccione un tipo de moneda", vbOKOnly + vbInformation
End Sub

Private Sub FiltrarRecordSet(Indice As Integer)
    DoEvents
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(2) As String
    Dim i As Integer
    If lrsTabla(Indice) Is Nothing Then Exit Sub
    If IsNull(dtpFechaBus) Then Exit Sub
    cadena = ""
    If chkFecha(Indice).Value = 1 Then filtros(2) = "Tca_dFecha like '" & Me.dtpFechaBus(Indice) & "'"
    For i = 0 To 2
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    ' *** Filtrando segun campos
    lrsTabla(Indice).Filter = 0
    If Trim(cadena) <> "" Then
        On Error Resume Next
        lrsTabla(Indice).Filter = cadena
    Else
        lrsTabla(Indice).Filter = 0
    End If
End Sub

Private Sub tdbcMes_ItemChange(Index As Integer)
    ConfigurarControlFecha (Index)
    CargaTablaMoneda (Index)
End Sub

Private Sub tdbcTipoMensual_ItemChange()
    Call CargaMeses
End Sub

Private Sub tdbgMoneda_GotFocus(Index As Integer)
    On Error Resume Next
    tdbgMoneda(IndiceTab).HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgMoneda_HeadClick(Index As Integer, ByVal ColIndex As Integer)
If Not lrsTabla(Index) Is Nothing Then
    If lrsTabla(Index).RecordCount > 0 Then
        lrsTabla(Index).Sort = tdbgMoneda(Index).Columns(ColIndex).DataField
        tdbgMoneda(Index).DataSource = lrsTabla(Index)
    End If
End If
End Sub

Private Sub tdbgMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Editar
    End If
End Sub

Private Sub tdbgMoneda_LostFocus(Index As Integer)
    tdbgMoneda(Index).HighlightRowStyle = ""
End Sub

Private Sub tdbcMonedaAdic_ItemChange()
    If IndiceTab < 2 Then
    tdbcMes_ItemChange (IndiceTab)
    End If

End Sub

Private Sub tdbnVentaP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grabar
    End If
End Sub


