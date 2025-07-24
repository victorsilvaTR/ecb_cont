VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   Icon            =   "frmInportacionXLS1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8550
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTodo 
      Height          =   8475
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   11595
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInportacionXLS1.frx":0ECA
         Left            =   5880
         List            =   "frmInportacionXLS1.frx":0EE6
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   600
         Width           =   5055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Seleccion de Campos Entidad"
         Height          =   4560
         Index           =   0
         Left            =   255
         TabIndex        =   68
         Top             =   1935
         Width           =   10800
         Begin VB.TextBox txtETipoEntidad 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   87
            Top             =   400
            Width           =   855
         End
         Begin VB.TextBox txtERazonSocial 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   86
            Top             =   885
            Width           =   855
         End
         Begin VB.TextBox txtETipoPersona 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   85
            ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txtETipoDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   84
            ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox txtERuc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   83
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox txtEDireccion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   82
            Top             =   1440
            Width           =   855
         End
         Begin VB.Frame Frame1 
            Caption         =   "Seleccion de Campos Tipo Cambio"
            Height          =   4560
            Index           =   1
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   10800
            Begin VB.TextBox txtTMonedaDestino 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   75
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox txtTTCCompra 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   74
               Top             =   2040
               Width           =   855
            End
            Begin VB.TextBox txtTTCVenta 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   73
               ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox txtTTCVentaP 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   72
               ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
               Top             =   3240
               Width           =   855
            End
            Begin VB.TextBox txtTMonedaOrigen 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   71
               Top             =   885
               Width           =   855
            End
            Begin VB.TextBox txtTFechaCambio 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   70
               Top             =   400
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "T.C. Venta P"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   3360
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "T.C. Venta"
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   2760
               Width           =   855
            End
            Begin VB.Label Label10 
               Caption         =   "T.C. Compra"
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Moneda Destino"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   1500
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "Moneda Origen"
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label13 
               Caption         =   "Fecha Cambio"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   480
               Width           =   1215
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Entidad"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Razon Social"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Direccion"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Ruc"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Doc."
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo Persona."
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   3360
            Width           =   1215
         End
      End
      Begin TDBText6Ctl.TDBText tdbtArchivo 
         Height          =   375
         Left            =   180
         TabIndex        =   95
         Top             =   1395
         Width           =   10710
         _Version        =   65536
         _ExtentX        =   18891
         _ExtentY        =   661
         Caption         =   "frmInportacionXLS1.frx":0F64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInportacionXLS1.frx":0FD0
         Key             =   "frmInportacionXLS1.frx":0FEE
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
         Appearance      =   0
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
      Begin MSComctlLib.ProgressBar pbAvance 
         Height          =   195
         Left            =   180
         TabIndex        =   96
         Top             =   7680
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TDBText6Ctl.TDBText lblCorrelativo 
         Height          =   375
         Left            =   180
         TabIndex        =   97
         Top             =   540
         Width           =   4830
         _Version        =   65536
         _ExtentX        =   8520
         _ExtentY        =   661
         Caption         =   "frmInportacionXLS1.frx":1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInportacionXLS1.frx":109E
         Key             =   "frmInportacionXLS1.frx":10BC
         BackColor       =   14737632
         EditMode        =   0
         ForeColor       =   16711680
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   1
         AlignHorizontal =   2
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
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   5535
         TabIndex        =   108
         Top             =   7995
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImportarDatos 
         Height          =   435
         Left            =   1995
         TabIndex        =   107
         Top             =   7980
         Width           =   1665
         Caption         =   " Importar Datos"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   3780
         TabIndex        =   106
         Top             =   7995
         Width           =   1665
         Caption         =   " Imprimir Errores"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   5040
         TabIndex        =   105
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   540
         Width           =   450
         PicturePosition =   262148
         Size            =   "794;688"
         Picture         =   "frmInportacionXLS1.frx":1100
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CORRELATIVO:"
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
         Index           =   3
         Left            =   180
         TabIndex        =   104
         Top             =   225
         Width           =   1275
      End
      Begin MSForms.CommandButton cmdSeleccionar 
         Height          =   390
         Left            =   10920
         TabIndex        =   103
         Top             =   1350
         Width           =   450
         PicturePosition =   262148
         Size            =   "794;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PROCESO:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   102
         Top             =   7320
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE LOS DATOS A IMPORTAR:"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   101
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label lblAvance 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
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
         Left            =   -11100
         TabIndex        =   100
         Top             =   6240
         Width           =   5325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE ARCHIVO:"
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
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   99
         Top             =   1080
         Width           =   1950
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   435
         Left            =   240
         TabIndex        =   98
         Top             =   7980
         Width           =   1665
         Caption         =   "Cargar Configuracion"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccion de Campos Registro de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   11160
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   33
         Top             =   400
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   32
         Top             =   885
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   31
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   2825
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   30
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   2340
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1855
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   28
         Top             =   1370
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   27
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   3310
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   26
         Top             =   3795
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   25
         Top             =   4280
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   24
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   1370
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   23
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   885
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   22
         Top             =   400
         Width           =   495
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   21
         Top             =   4765
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   20
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   1855
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2340
         Width           =   495
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2825
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   17
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   4765
         Width           =   495
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   16
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   4280
         Width           =   495
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   15
         Top             =   3795
         Width           =   495
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   14
         Top             =   3310
         Width           =   495
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   13
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   1335
         Width           =   495
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   1815
         Width           =   495
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2295
         Width           =   495
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2790
         Width           =   495
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   7
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   4725
         Width           =   495
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   6
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   4245
         Width           =   495
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3750
         Width           =   495
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   4
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   3
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   1335
         Width           =   495
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   2
         ToolTipText     =   "01 - DNI , 02- PART NAC, 03-CARNET EXTRAN.,04-R.U.C ,05-PASAPORTE , 06-DOCVARIOS"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Moneda Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Ejercicio/Año"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Libro"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Glosa Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Debe Soles"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Haber Soles"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Tipo Cambio"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Debe Extranjera "
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Haber Extranjera "
         Height          =   255
         Left            =   2880
         TabIndex        =   55
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Left            =   2880
         TabIndex        =   54
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Centro de Costo"
         Height          =   255
         Left            =   2880
         TabIndex        =   53
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Tipo Entidad"
         Height          =   255
         Left            =   2880
         TabIndex        =   52
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Ruc Entidad"
         Height          =   255
         Left            =   2880
         TabIndex        =   51
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Tipo Documento"
         Height          =   255
         Left            =   2880
         TabIndex        =   50
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Fecha Documento"
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label32 
         Caption         =   "Serie Documento"
         Height          =   255
         Left            =   2880
         TabIndex        =   48
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "Numero Doc."
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "Fecha Venci."
         Height          =   255
         Left            =   5760
         TabIndex        =   46
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Documento Ref."
         Height          =   255
         Left            =   5760
         TabIndex        =   45
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Fecha Referen."
         Height          =   255
         Left            =   5760
         TabIndex        =   44
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "Serie Referencia"
         Height          =   255
         Left            =   5760
         TabIndex        =   43
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Numero Referen."
         Height          =   255
         Left            =   5760
         TabIndex        =   42
         Top             =   2355
         Width           =   1215
      End
      Begin VB.Label Label39 
         Caption         =   "Monto Inafecto"
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   2835
         Width           =   1215
      End
      Begin VB.Label Label40 
         Caption         =   "Retencion"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   3315
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "Fecha Retencion"
         Height          =   255
         Left            =   5760
         TabIndex        =   39
         Top             =   3795
         Width           =   1335
      End
      Begin VB.Label Label42 
         Caption         =   "Numero Retencion"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   4275
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Tipo de Provision."
         Height          =   255
         Left            =   5760
         TabIndex        =   37
         Top             =   4755
         Width           =   1455
      End
      Begin VB.Label Label44 
         Caption         =   "Tipo de Oper. TC."
         Height          =   255
         Left            =   8400
         TabIndex        =   36
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label Label45 
         Caption         =   "Tipo Moneda"
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label46 
         Caption         =   "Com.No Domiciliado"
         Height          =   255
         Left            =   8400
         TabIndex        =   34
         Top             =   1395
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
