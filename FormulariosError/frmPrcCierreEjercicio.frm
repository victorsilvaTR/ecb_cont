VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcCierreEjercicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Cierre del Ejercicio"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   Icon            =   "frmPrcCierreEjercicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6390
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5655
      Left            =   6480
      TabIndex        =   29
      Top             =   315
      Visible         =   0   'False
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Generar Asientos de Cierre del Ejercicio"
      TabPicture(0)   =   "frmPrcCierreEjercicio.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Hoja de Trabajo de Impuesto a la Renta"
      TabPicture(1)   =   "frmPrcCierreEjercicio.frx":0EE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5235
         Left            =   135
         TabIndex        =   37
         Top             =   360
         Width           =   9780
         Begin VB.Frame Frame2 
            BackColor       =   &H80000009&
            ForeColor       =   &H00C00000&
            Height          =   3345
            Left            =   135
            TabIndex        =   42
            Top             =   1080
            Width           =   9555
            Begin VB.TextBox tdbHTtc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   57
               Text            =   "210111.02"
               Top             =   2340
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.TextBox tdbHTUtilDist 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   25
               Text            =   "0.00"
               Top             =   1980
               Width           =   1965
            End
            Begin VB.TextBox tdbHTReservaLegal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   24
               Text            =   "0.00"
               Top             =   1620
               Width           =   1965
            End
            Begin VB.TextBox tdbHTUtilNetaEjer 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   23
               Text            =   "0.00"
               Top             =   1260
               Width           =   1965
            End
            Begin VB.TextBox tdbHTImpRenta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   22
               Text            =   "0.00"
               Top             =   900
               Width           =   1965
            End
            Begin VB.TextBox tdbHTRentaNetaImponible 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "0.00"
               Top             =   540
               Width           =   1965
            End
            Begin VB.TextBox tdbHTParticip 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   "0.00"
               Top             =   180
               Width           =   1965
            End
            Begin VB.TextBox tdbHTRentaNetaAntesImp 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "0.00"
               Top             =   2790
               Width           =   1965
            End
            Begin VB.TextBox tdbHTPerdidaNeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2610
               TabIndex        =   18
               Text            =   "0.00"
               Top             =   2340
               Width           =   1965
            End
            Begin VB.TextBox tdbHTPerdidaEj 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0.00"
               Top             =   1980
               Width           =   1965
            End
            Begin VB.TextBox tdbHTRentaEj 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0.00"
               Top             =   1620
               Width           =   1965
            End
            Begin VB.TextBox tdbHTDeducc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2610
               TabIndex        =   15
               Text            =   "0.00"
               Top             =   1260
               Width           =   1965
            End
            Begin VB.TextBox tdbHTAdic 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2610
               TabIndex        =   14
               Text            =   "0.00"
               Top             =   900
               Width           =   1965
            End
            Begin VB.TextBox tdbHTPerdCont 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "0.00"
               Top             =   540
               Width           =   1965
            End
            Begin VB.TextBox tdbHTUtilcont 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "210111.02"
               Top             =   180
               Width           =   1965
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Utilidad distribuible"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   4905
               TabIndex        =   56
               Top             =   1980
               Width           =   1635
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Reserva legal                 (%)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   4905
               TabIndex        =   55
               Top             =   1620
               Width           =   2460
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Utilidad neta del ejercicio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   4905
               TabIndex        =   54
               Top             =   1260
               Width           =   2190
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Impuesto a la renta         (%)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   4905
               TabIndex        =   53
               Top             =   900
               Width           =   2445
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Renta neta imponible"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   4905
               TabIndex        =   52
               Top             =   540
               Width           =   1815
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Particip. de los Trab.       (%)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   4905
               TabIndex        =   51
               Top             =   225
               Width           =   2475
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Renta neta del ejer. antes de impuestos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   7
               Left            =   180
               TabIndex        =   50
               Top             =   2835
               Width           =   2370
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pérdida neta compensable de ejercicios anteriores"
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
               Height          =   390
               Index           =   6
               Left            =   180
               TabIndex        =   49
               Top             =   2295
               Width           =   2325
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pérdida del ejercicio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   180
               TabIndex        =   48
               Top             =   1980
               Width           =   1755
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Utilidad del ejercicio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   47
               Top             =   1620
               Width           =   1755
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "(-)  Deducciones"
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
               Index           =   3
               Left            =   180
               TabIndex        =   46
               Top             =   1260
               Width           =   1425
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "(+) Adiciones"
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
               Left            =   180
               TabIndex        =   45
               Top             =   900
               Width           =   1125
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Pérdida contable"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   44
               Top             =   540
               Width           =   1455
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Utilidad Contable"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   43
               Top             =   225
               Width           =   1470
            End
         End
         Begin TDBText6Ctl.TDBText tdbtPorcPartic 
            Height          =   285
            Left            =   2970
            TabIndex        =   7
            Top             =   630
            Width           =   420
            _Version        =   65536
            _ExtentX        =   741
            _ExtentY        =   503
            Caption         =   "frmPrcCierreEjercicio.frx":0F02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcCierreEjercicio.frx":0F6E
            Key             =   "frmPrcCierreEjercicio.frx":0F8C
            BackColor       =   16777152
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
            MaxLength       =   3
            LengthAsByte    =   0
            Text            =   "5"
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
         Begin TDBText6Ctl.TDBText tdbtPorcRenta 
            Height          =   285
            Left            =   5940
            TabIndex        =   9
            Top             =   630
            Width           =   420
            _Version        =   65536
            _ExtentX        =   741
            _ExtentY        =   503
            Caption         =   "frmPrcCierreEjercicio.frx":0FD0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcCierreEjercicio.frx":103C
            Key             =   "frmPrcCierreEjercicio.frx":105A
            BackColor       =   16777152
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
            MaxLength       =   3
            LengthAsByte    =   0
            Text            =   "30"
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
         Begin TDBText6Ctl.TDBText tdbtPorcReserva 
            Height          =   285
            Left            =   8820
            TabIndex        =   11
            Top             =   630
            Width           =   420
            _Version        =   65536
            _ExtentX        =   741
            _ExtentY        =   503
            Caption         =   "frmPrcCierreEjercicio.frx":109E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcCierreEjercicio.frx":110A
            Key             =   "frmPrcCierreEjercicio.frx":1128
            BackColor       =   16777152
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
            MaxLength       =   3
            LengthAsByte    =   0
            Text            =   "5"
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
         Begin MSForms.CommandButton cmdImprimir 
            Height          =   435
            Left            =   7065
            TabIndex        =   60
            ToolTipText     =   "Imprimir la hora de trabajo"
            Top             =   4590
            Width           =   1260
            Caption         =   " Imprimir"
            PicturePosition =   327683
            Size            =   "2222;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdSalir 
            Height          =   435
            Index           =   0
            Left            =   8370
            TabIndex        =   59
            ToolTipText     =   " Salir de cierre del ejercicio"
            Top             =   4590
            Width           =   1260
            Caption         =   " Salir"
            PicturePosition =   327683
            Size            =   "2222;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdLeerDatos 
            Height          =   435
            Left            =   135
            TabIndex        =   58
            ToolTipText     =   " Carga los datos guardados de la hoja de trabajo"
            Top             =   4590
            Width           =   1470
            Caption         =   " Cargar Datos"
            PicturePosition =   327683
            Size            =   "2593;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdGrabarDatos 
            Height          =   435
            Left            =   1665
            TabIndex        =   27
            ToolTipText     =   "Graba los datos de la hoja de trabajo"
            Top             =   4590
            Width           =   1260
            VariousPropertyBits=   25
            Caption         =   " Grabar"
            PicturePosition =   327683
            Size            =   "2222;767"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdCargarDatos 
            Height          =   435
            Left            =   3780
            TabIndex        =   26
            ToolTipText     =   "Calcula los importes"
            Top             =   4590
            Width           =   1260
            Caption         =   " Calcular"
            PicturePosition =   327683
            Size            =   "2222;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdGenerarImpuesto 
            Height          =   435
            Left            =   5220
            TabIndex        =   28
            ToolTipText     =   "Genera los asientos del ejercicio"
            Top             =   4590
            Width           =   1260
            VariousPropertyBits=   25
            Caption         =   " Generar"
            PicturePosition =   327683
            Size            =   "2222;767"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CheckBox chkReserva 
            Height          =   345
            Left            =   7290
            TabIndex        =   10
            Top             =   585
            Width           =   1440
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2540;609"
            Value           =   "0"
            Caption         =   "Reserva Legal"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label133 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   9315
            TabIndex        =   41
            Top             =   675
            Width           =   225
         End
         Begin VB.Label Label133 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   6435
            TabIndex        =   40
            Top             =   675
            Width           =   225
         End
         Begin VB.Label Label133 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3420
            TabIndex        =   39
            Top             =   675
            Width           =   225
         End
         Begin MSForms.CheckBox chkPorcImp 
            Height          =   345
            Left            =   4050
            TabIndex        =   8
            Top             =   585
            Width           =   1830
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3228;609"
            Value           =   "0"
            Caption         =   "Impuesto a la Renta"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPorcParticip 
            Height          =   345
            Left            =   315
            TabIndex        =   6
            Top             =   585
            Width           =   2685
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4736;609"
            Value           =   "0"
            Caption         =   "Participación de los trabajadores"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "HOJA DE TRABAJO DEL IMPUESTO A LA RENTA"
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
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   225
            Width           =   9510
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4275
         Left            =   -74640
         TabIndex        =   30
         Top             =   465
         Width           =   7740
         Begin VB.CheckBox chkTipo 
            Caption         =   "Por Tipo"
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
            Left            =   360
            TabIndex        =   31
            Top             =   600
            Width           =   1095
         End
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   2595
            Left            =   360
            TabIndex        =   32
            Top             =   1440
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   4577
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Tipo de Ratio"
            Columns(0).DataField=   "Ind_cTipoRatios"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre Tipo de Ratio"
            Columns(1).DataField=   "Tab_cDescripCampo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Indicador"
            Columns(2).DataField=   "Ind_cCodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Descripcion"
            Columns(3).DataField=   "Ind_cDescripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "% Min"
            Columns(4).DataField=   "Ind_nPorceMin"
            Columns(4).NumberFormat=   "Standard"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "% Max"
            Columns(5).DataField=   "Ind_nPorceMax"
            Columns(5).NumberFormat=   "Standard"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3016"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2937"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1296"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1217"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4207"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4128"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=528"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1323"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1244"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=528"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1164"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1085"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=532"
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
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HF1EFEB&"
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
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1560
            TabIndex        =   33
            Tag             =   "_"
            Top             =   960
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   556
            Caption         =   "frmPrcCierreEjercicio.frx":116C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcCierreEjercicio.frx":11D8
            Key             =   "frmPrcCierreEjercicio.frx":11F6
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
            AllowSpace      =   0
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   250
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
         Begin TrueOleDBList70.TDBCombo tdbcTipoBus 
            Height          =   300
            Left            =   1560
            TabIndex        =   34
            Tag             =   "_"
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
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
            _PropDict       =   $"frmPrcCierreEjercicio.frx":1248
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
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
            Index           =   14
            Left            =   360
            TabIndex        =   36
            Top             =   1000
            Width           =   990
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
            TabIndex        =   35
            Top             =   240
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraTodo 
      Height          =   4005
      Left            =   90
      TabIndex        =   61
      Top             =   45
      Width           =   6255
      Begin VB.Frame Frame3 
         Height          =   2820
         Left            =   315
         TabIndex        =   62
         Top             =   585
         Width           =   5715
         Begin VB.OptionButton optAsiento 
            Caption         =   "Asiento automático Extendido"
            Height          =   285
            Index           =   1
            Left            =   1530
            TabIndex        =   2
            Top             =   1665
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.OptionButton optAsiento 
            Caption         =   "Asiento automático Simplificado"
            Height          =   285
            Index           =   0
            Left            =   1530
            TabIndex        =   1
            Top             =   1305
            Value           =   -1  'True
            Width           =   2670
         End
         Begin TrueOleDBList70.TDBCombo tdbcLibro 
            Height          =   300
            Left            =   270
            TabIndex        =   0
            Tag             =   "enabled"
            Top             =   855
            Width           =   5055
            _ExtentX        =   8916
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
            _PropDict       =   $"frmPrcCierreEjercicio.frx":12CF
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
         Begin TDBDate6Ctl.TDBDate dtpFechaCierre 
            Height          =   300
            Left            =   3015
            TabIndex        =   3
            Top             =   2160
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   529
            Calendar        =   "frmPrcCierreEjercicio.frx":1356
            Caption         =   "frmPrcCierreEjercicio.frx":1458
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcCierreEjercicio.frx":14BC
            Keys            =   "frmPrcCierreEjercicio.frx":14DA
            Spin            =   "frmPrcCierreEjercicio.frx":1546
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2010185729
            Value           =   38202
            CenturyMode     =   0
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de cierre"
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
            Left            =   1080
            TabIndex        =   65
            Top             =   2205
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Seleccione el Libro donde se generarán los asientos automáticos, los asientos se crearan en el mes de AJUSTE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   1
            Left            =   270
            TabIndex        =   63
            Top             =   180
            Width           =   5040
         End
      End
      Begin MSForms.CommandButton cmdGenerarCierre 
         Height          =   435
         Left            =   1350
         TabIndex        =   4
         ToolTipText     =   "Genera los asientos de cierre"
         Top             =   3465
         Width           =   1665
         Caption         =   " Generar"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcCierreEjercicio.frx":156E
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Caption         =   "GENERAR ASIENTOS DE CIERRE DEL EJERCICIO"
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
         Left            =   315
         TabIndex        =   64
         Top             =   270
         Width           =   5730
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         ToolTipText     =   " Salir de cierre del ejercicio"
         Top             =   3465
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcCierreEjercicio.frx":1B08
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
      TabIndex        =   66
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcCierreEjercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lrsTabla As ADODB.Recordset
Dim gsGrupo As String
Dim lArrReg() As Variant

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub CargaArreglo()
    ReDim lArrReg(22) As Variant
    
    lArrReg(0) = "BUSCARREGISTRO"
    lArrReg(1) = gsEmpresa
    lArrReg(2) = gsAnio
    lArrReg(3) = NE(tdbtPorcPartic)
    lArrReg(4) = NE(tdbtPorcRenta)
    lArrReg(5) = NE(tdbtPorcReserva)
    
    lArrReg(6) = NE(tdbHTUtilcont)
    lArrReg(7) = NE(tdbHTPerdCont)
    lArrReg(8) = NE(tdbHTAdic)
    lArrReg(9) = NE(tdbHTDeducc)
    lArrReg(10) = NE(tdbHTRentaEj)
    lArrReg(11) = NE(tdbHTPerdidaEj)
    lArrReg(12) = NE(tdbHTPerdidaNeta)
    lArrReg(13) = NE(tdbHTRentaNetaAntesImp)
    lArrReg(14) = NE(tdbHTParticip)
    lArrReg(15) = NE(tdbHTRentaNetaImponible)
    lArrReg(16) = NE(tdbHTImpRenta)
    lArrReg(17) = NE(tdbHTUtilNetaEjer)
    lArrReg(18) = NE(tdbHTReservaLegal)
    lArrReg(19) = NE(tdbHTUtilDist)
    
    If chkPorcParticip.Value = True Then
        lArrReg(20) = 1
    Else
        lArrReg(20) = 0
    End If
    
    If chkPorcImp.Value = True Then
        lArrReg(21) = 1
    Else
        lArrReg(21) = 0
    End If
    
    If chkReserva.Value = True Then
        lArrReg(22) = 1
    Else
        lArrReg(22) = 0
    End If
    
End Sub

Private Sub LimpiarCampos()
    tdbtPorcPartic = "0"
    tdbtPorcRenta = "0"
    tdbtPorcReserva = "0"
    
    tdbHTUtilcont = "0.00"
    tdbHTPerdCont = "0.00"
    tdbHTAdic = "0.00"
    tdbHTDeducc = "0.00"
    tdbHTRentaEj = "0.00"
    tdbHTPerdidaEj = "0.00"
    tdbHTPerdidaNeta = "0.00"
    tdbHTRentaNetaAntesImp = "0.00"
    tdbHTParticip = "0.00"
    tdbHTRentaNetaImponible = "0.00"
    tdbHTImpRenta = "0.00"
    tdbHTUtilNetaEjer = "0.00"
    tdbHTReservaLegal = "0.00"
    tdbHTUtilDist = "0.00"
    
    chkPorcParticip.Value = vbUnchecked
    chkPorcImp.Value = vbUnchecked
    chkReserva.Value = vbUnchecked

End Sub

Private Function BuscaHojaTrabajo() As Boolean
    Dim clsMante As New clsMantoTablas
    Dim sqlver As String
    Dim rsArreglo  As New ADODB.Recordset
    Dim arrDatos() As Variant
    
    
    Screen.MousePointer = vbHourglass
    
    sqlver = "spCNT_HOJA_TRABAJO 'BUSCARREGISTRO', '" & gsEmpresa & "', '" & gsAnio & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
    arrDatos = Array(sqlver)
    Set rsArreglo = clsMante.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = adStateOpen Then
        If Not rsArreglo.EOF And Not rsArreglo.BOF Then
            Do While Not rsArreglo.EOF
           
              tdbtPorcPartic = NE(rsArreglo!Porc_Particip)
              tdbtPorcRenta = NE(rsArreglo!Porc_Renta)
              tdbtPorcReserva = NE(rsArreglo!Porc_Reserva)
              
              tdbHTUtilcont = Format(NE(rsArreglo!Ej_nUtilCont), "###,###,##0.00")
              tdbHTPerdCont = Format(Format(NE(rsArreglo!Ej_nPerdCont)), "###,###,##0.00")
              tdbHTAdic = Format(NE(rsArreglo!Ej_nAdiciones), "###,###,##0.00")
              tdbHTDeducc = Format(NE(rsArreglo!Ej_nDeducciones), "###,###,##0.00")
              tdbHTRentaEj = Format(NE(rsArreglo!Ej_nUtilEjercicio), "###,###,##0.00")
              tdbHTPerdidaEj = Format(NE(rsArreglo!Ej_nPerdEjercicio), "###,###,##0.00")
              tdbHTPerdidaNeta = Format(NE(rsArreglo!Ej_nPerdNeta), "###,###,##0.00")
              tdbHTRentaNetaAntesImp = Format(NE(rsArreglo!Ej_nRentaAntesImp), "###,###,##0.00")
              tdbHTParticip = Format(NE(rsArreglo!Ej_nParticipaciones), "###,###,##0.00")
              tdbHTRentaNetaImponible = Format(NE(rsArreglo!Ej_nRentaNeta), "###,###,##0.00")
              tdbHTImpRenta = Format(NE(rsArreglo!Ej_nImpRenta), "###,###,##0.00")
              tdbHTUtilNetaEjer = Format(NE(rsArreglo!Ej_nUtilNeta), "###,###,##0.00")
              tdbHTReservaLegal = Format(NE(rsArreglo!Ej_nReserva), "###,###,##0.00")
              tdbHTUtilDist = Format(NE(rsArreglo!Ej_nUtilDist), "###,###,##0.00")
              
              
                If NE(rsArreglo!Chk_Partic) = 1 Then
                    chkPorcParticip.Value = vbChecked
                Else
                    chkPorcParticip.Value = vbUnchecked
                End If
                
                If NE(rsArreglo!Chk_Renta) = 1 Then
                    chkPorcImp.Value = vbChecked
                Else
                    chkPorcImp.Value = vbUnchecked
                End If
                
                If NE(rsArreglo!Chk_Reserva) = 1 Then
                    chkReserva.Value = vbChecked
                Else
                    chkReserva.Value = vbUnchecked
                End If
              
              
              rsArreglo.MoveNext
            Loop
        Else
            LimpiarCampos
        End If
    Else
        LimpiarCampos
    End If
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
    
End Function


Private Function GrabaHojaTrabajo() As Boolean
    Dim clsMante As clsMantoTablas
    Dim sqlver As String
    
    Set clsMante = New clsMantoTablas
    
    Call CargaArreglo
    
    Screen.MousePointer = vbHourglass
    
    lArrReg(0) = "INSERTAR"
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCNT_HOJA_TRABAJO", lArrReg(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
          
        Set clsMante = Nothing
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    Mensajes "Se grabo la hoja de trabajo", vbInformation + vbOKOnly
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
End Function

Private Sub chkPorcImp_Click()
    If ValidaCheck = False Then Exit Sub
    
    If chkPorcImp.Value = True Then
        'chkPorcParticip.Value = True
    Else
        'chkReserva.Value = False
    End If


    CalculaProcentajes
End Sub

Private Function ValidaCheck() As Boolean
    ValidaCheck = True
    If NE(tdbHTPerdidaEj) > 0 Then
        chkPorcParticip.Value = False
        chkPorcImp.Value = False
        chkReserva.Value = False
        ValidaCheck = False
    End If
    
End Function

Private Sub chkPorcParticip_Click()
    
    If ValidaCheck = False Then Exit Sub
    
    If chkPorcParticip.Value = False Then
        'chkPorcImp.Value = False
        'chkReserva.Value = False
    End If

    CalculaProcentajes
End Sub

Private Sub CalculaProcentajes()

    If chkPorcParticip.Value = True Then
         ActivarControlPersup tdbHTParticip, True
         tdbHTParticip = NE(tdbHTRentaNetaAntesImp) * NE(tdbtPorcPartic) / 100
         
         ActivarControlPersup tdbHTRentaNetaImponible, True
         tdbHTRentaNetaImponible = NE(tdbHTRentaNetaAntesImp) - NE(tdbHTParticip)
         
    Else
         ActivarControlPersup tdbHTParticip, False
         tdbHTParticip = "0.00"
        
         ActivarControlPersup tdbHTRentaNetaImponible, True
         tdbHTRentaNetaImponible = NE(tdbHTRentaNetaAntesImp) - NE(tdbHTParticip)
        
    End If

    If chkPorcImp.Value = True Then
         ActivarControlPersup tdbHTImpRenta, True
         tdbHTImpRenta = NE(tdbHTRentaNetaImponible) * NE(tdbtPorcRenta) / 100
         
         ActivarControlPersup tdbHTUtilNetaEjer, True
         tdbHTUtilNetaEjer = NE(tdbHTRentaNetaImponible) - NE(tdbHTImpRenta)
    Else
         ActivarControlPersup tdbHTImpRenta, False
         tdbHTImpRenta = "0.00"
        
         ActivarControlPersup tdbHTUtilNetaEjer, True
         tdbHTUtilNetaEjer = NE(tdbHTRentaNetaImponible) - NE(tdbHTImpRenta)
    End If

    If chkReserva.Value = True Then
         ActivarControlPersup tdbHTReservaLegal, True
         tdbHTReservaLegal = NE(tdbHTUtilNetaEjer) * NE(tdbtPorcReserva) / 100
         
         ActivarControlPersup tdbHTUtilDist, True
         tdbHTUtilDist = NE(tdbHTUtilNetaEjer) - NE(tdbHTReservaLegal)
    Else
        ActivarControlPersup tdbHTReservaLegal, False
        tdbHTReservaLegal = "0.00"
        
        ActivarControlPersup tdbHTUtilDist, True
        tdbHTUtilDist = NE(tdbHTUtilNetaEjer) - NE(tdbHTReservaLegal)
    End If

    ControlAbs tdbHTParticip
    ControlAbs tdbHTRentaNetaImponible
    ControlAbs tdbHTImpRenta
    ControlAbs tdbHTUtilNetaEjer
    ControlAbs tdbHTReservaLegal
    ControlAbs tdbHTUtilDist

End Sub

Private Sub chkReserva_Click()
    If ValidaCheck = False Then Exit Sub

    If chkReserva.Value = True Then
        'chkPorcParticip.Value = True
        'chkPorcImp.Value = True
    End If

    CalculaProcentajes
End Sub

Private Sub CargaUtilidadPerdidaContable()
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    tdbHTUtilcont = "0.00"
    tdbHTPerdCont = "0.00"
    
    sqlver = "SELECT PLA_CCUENTACONTABLE, " & _
             "(SUM(Asd_nDebeSoles - Asd_nHaberSoles)) as MontoSoles, " & _
             "(SUM(Asd_nDebeMonExt - Asd_nHaberMonExt)) as MontoDolares " & _
             "From CND_ASIENTO_VOUCHER " & _
             "WHERE EMP_CCODIGO = '" & gsEmpresa & "' AND " & _
             "PAN_CANIO = '" & gsAnio & "' AND " & _
             "PER_CPERIODO = '13' AND " & _
             "ASD_CDELETED <>'*' AND " & _
             "left(PLA_CCUENTACONTABLE,2) ='89' " & _
             "GROUP BY PLA_CCUENTACONTABLE "

    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
    If rsArreglo.State = adStateOpen Then
    
       If NE(rsArreglo("MontoSoles")) > 0 Then
            tdbHTPerdCont = NE(rsArreglo("MontoSoles"))
       Else
            tdbHTUtilcont = NE(rsArreglo("MontoSoles"))
       End If
       
       
       
       
       Dim TC As Double
       TC = NE(rsArreglo("MontoDolares"))
       If TC <= 0 Then
            TC = 1
       Else
          TC = NE(rsArreglo("MontoSoles")) / NE(rsArreglo("MontoDolares"))
       End If
       
       If gsByMoneda = 0 Then TC = 1
       
       tdbHTtc = TC
       
       ControlAbs tdbHTUtilcont
       ControlAbs tdbHTPerdCont
       ControlAbs tdbHTtc
       
    End If
    End If
    
    Call CerrarRecordSet(rsArreglo)

    
End Sub

Private Sub cmdCargarDatos_Click()
    CargaUtilidadPerdidaContable
    CalculaRentayPerdidaEjercicio
End Sub

Private Function EliminarAsientosCierre()
Screen.MousePointer = vbHourglass
Dim clsMante As clsMantoTablas
Dim lArrMnt(10) As Variant
Set clsMante = New clsMantoTablas
lArrMnt(0) = "ELIMINACIERRE"
lArrMnt(1) = ""
lArrMnt(2) = gsEmpresa
lArrMnt(3) = gsAnio
lArrMnt(4) = "13"
lArrMnt(5) = tdbcLibro.BoundText
lArrMnt(6) = ""
lArrMnt(7) = ""
lArrMnt(8) = ""
lArrMnt(9) = ""
lArrMnt(10) = ""

If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ConsultaAsientos", lArrMnt(), True) = False Then
    Mensajes "El proceso no ha concluido. Verificar...", vbInformation
    Screen.MousePointer = vbDefault
    'EliminarAsientosCierre = False
    Set clsMante = Nothing
    Exit Function
End If

Set clsMante = Nothing

Screen.MousePointer = vbDefault

End Function



Private Function BuscaCuentasCierre() As Boolean
    Dim sql As String
    Dim cadena As String
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    BuscaCuentasCierre = True
    
    If optAsiento(0).Value = True Then
        cadena = "BUSCA_CTARESULT_OP"
    Else
        cadena = "BUSCA_CTASCIERRE_OP"
    End If
    
    
    sql = "spCn_ConsultaCuentas '" & cadena & "','" & gsEmpresa & "','" & gsAnio & "'"
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    cadena = ""
    If Not rsAddItem Is Nothing Then
        If rsAddItem.State = adStateOpen Then
            Do While Not rsAddItem.EOF
                cadena = cadena & rsAddItem.AbsolutePosition & ") " & CE(rsAddItem!COP_CDESCRIPCION) & Salto(1)
                BuscaCuentasCierre = False
                rsAddItem.MoveNext
            Loop
        End If
    End If
    
    If BuscaCuentasCierre = False Then
        If optAsiento(0).Value = True Then
            Mensajes "En el Plan de cuentas, falta ingresar la cuenta de cierre de :" & Salto(2) & cadena & Salto(1) & "Primero configure las cuentas y ejecute nuevamente este proceso ...", vbOKOnly + vbInformation
        Else
            Mensajes "En el Plan de cuentas, faltan ingresar las cuentas de cierre de :" & Salto(2) & cadena & Salto(1) & "Primero configure las cuentas y ejecute nuevamente este proceso ...", vbOKOnly + vbInformation
        End If
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing

End Function

Public Function GenerarCierreDolares(ByVal strCuenta As String)
    
    Dim sp As String
    Dim clsMante As clsMantoTablas
    Dim ArrayList(9) As Variant
    
    ArrayList(0) = gsEmpresa
    ArrayList(1) = gsAnio
    ArrayList(2) = "13" 'Me.tdbcLibro.BoundText
    ArrayList(3) = "040"
    ArrayList(4) = strCuenta
    ArrayList(5) = gsUsuario
    
    Set clsMante = New clsMantoTablas
    If gintBiMoneda = 1 Then
        'Moneda Extranjera
        sp = "spCn_CierreAnual_Dolares"
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, ArrayList(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
        End If
    End If
    
End Function

Private Sub cmdGenerarCierre_Click()
    Dim respuesta As String
    Dim sp As String
    Dim sql As String
    
    If CE(tdbcLibro.Text) = "" Then
        Mensajes "Cree el libro de cierre y configurelo en Parametros iniciales", vbOKOnly + vbInformation
        cmdGenerarCierre.Enabled = True
        cmdGenerarCierre.Enabled = False
        Exit Sub
    End If
   
    If IsDate(dtpFechaCierre.Value) = False Or dtpFechaCierre.Text = "__/__/____" Then
        Mensajes "Ingrese una Fecha de Cierre válida"
        cmdGenerarCierre.Enabled = True
        pSetFocus dtpFechaCierre
        Exit Sub
    End If
   
    If IsDate(dtpFechaCierre.Value) = True Then
        If dtpFechaCierre.Year <> gsAnio Then
        Mensajes "El año de la Fecha de Cierre, es invalido" & Salto(1) & "Debe ser igual al año del sistema " & gsAnio
        cmdGenerarCierre.Enabled = True
        pSetFocus dtpFechaCierre
        Exit Sub
        End If
    End If
   
    If CierreMes("13") Then
        Mensajes "El mes de AJUSTE esta bloqueado, no se puede generar los asientos automaticos"
        Exit Sub
    End If
    
    cmdGenerarCierre.Enabled = False
    
    If BuscaCuentasCierre = False Then
        cmdGenerarCierre.Enabled = True
        Exit Sub
    End If

    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '08' and Per_cPeriodo = '13' and Cic_cEstado = 'I'"
    If ExisteDato(sql) = True Then Mensajes "No se puede procesar el Ajuste por Cierre, debido a que el periodo se encuentra bloqueado", vbInformation: Exit Sub
    
    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '08' and Per_cPeriodo = '13'"
    If ExisteDato(sql) = True Then
        Mensajes "Esta corrección modificará los datos ingresados, la misma que será informada a la SUNAT en el período " + UCase(MonthName(Month(lsFecha))) + " del ejercicio " + Str(Year(lsFecha)) + "."
        If MsgBox("Desea continuar..?", vbQuestion + vbOKCancel, gsNombreModulo) = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    If optAsiento(0).Value = True Then
        respuesta = MsgBox(" " & Salto(2) & "Desea Generar el asiento de cierre del Año Actual" & Salto(1) & "En el periodo de AJUSTE", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Cierre")
        sp = "spCn_GeneraAsientoEjer_Simple"
    Else
        respuesta = MsgBox("MODO EXTENDIDO:" & Salto(2) & "Desea Generar los asientos de cierre del Año Actual" & Salto(1) & "En el periodo de AJUSTE", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Cierre")
        sp = "spCn_GeneraAsientoEjercicio"
    End If
    
    If respuesta = vbYes Then
        EliminarAsientosCierre
        DoEvents
        Call ActualizaSaldos
        DoEvents

        Screen.MousePointer = vbHourglass
        Dim clsMante As clsMantoTablas
        Dim lArrMnt(9) As Variant
        Set clsMante = New clsMantoTablas
        
        If gsTipoPlan = "1" Then
        
         sp = "spCn_CierreAnual"
        
        lArrMnt(0) = gsEmpresa
        lArrMnt(1) = gsAnio
        lArrMnt(2) = "13" 'Me.tdbcLibro.BoundText
        lArrMnt(3) = gsMonedaNac
'        lArrMnt(5) = "60"
        lArrMnt(5) = gsUsuario
        
        
        lArrMnt(4) = "61" 'Cuenta 61 a la 80
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("61")
        
        lArrMnt(4) = "71" 'Cuenta 71 a la 811
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("71")
        
        lArrMnt(4) = "715" 'Cuenta 715 a la 812
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("715")
        
        lArrMnt(4) = "79" 'Cuenta 79 Y 78 a la 80
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("79")
        
        lArrMnt(4) = "69" 'Cuenta 69 a la 61
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
               
        GenerarCierreDolares ("69")
        
        lArrMnt(4) = "70" 'Cuenta 70 a la 81
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("70")
        
        lArrMnt(4) = "701" 'Cuenta 701 a la 80
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("701")
        
        lArrMnt(4) = "74" 'Cuenta 74 a la 80
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If

        GenerarCierreDolares ("74")
        
        lArrMnt(4) = "60" 'Cuenta 60 a la 80
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        
'        lArrMnt(4) = "63" 'Cuenta 63 a la 82
'        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
'            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'            Screen.MousePointer = vbDefault
'            cmdGenerarCierre.Enabled = True
'            cmdSalir(1).Enabled = True
'            Exit Sub
'        End If
        GenerarCierreDolares ("60")
        
        lArrMnt(4) = "602" 'Cuenta 602 a la 82
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("602")
        
        lArrMnt(4) = "80" 'Cuenta 80 y 81 a la 82
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("80")
        
        lArrMnt(4) = "82" 'Cuenta 82 a la 83
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
    
        GenerarCierreDolares ("82")
        
        lArrMnt(4) = "62" 'Cuenta 62 y 64 a la 83
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("62")
        
        lArrMnt(4) = "83" 'Cuenta 83 a la 84
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("83")
        
        lArrMnt(4) = "813" 'Cuenta 813 a la 72
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("813")
        
        lArrMnt(4) = "75" 'Cuenta 75, 73 y 76 a la 841
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("75")
        
        lArrMnt(4) = "65" 'Cuenta 65 y 68 a la 84
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
      
        

'        lArrMnt(4) = "613" 'Cuenta 602 a la 82
'        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
'            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'            Screen.MousePointer = vbDefault
'            cmdGenerarCierre.Enabled = True
'            cmdSalir(1).Enabled = True
'            Exit Sub
'        End If
        

        GenerarCierreDolares ("65")

        lArrMnt(4) = "84" 'Cuenta 84 a la 85
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
           
        GenerarCierreDolares ("84")
        
        lArrMnt(4) = "77" 'Cuenta 77 a la 85
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("77")
        
        lArrMnt(4) = "7591" 'Cuenta 7591 a la 83
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("7591")
        
        lArrMnt(4) = "655" 'Cuenta 655,6591 y 6592 a la 85
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("655")
        
        lArrMnt(4) = "67" 'Cuenta 67 a la 85
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("67")
        
        lArrMnt(4) = "85"        'Cuenta 85 a la 89
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
        
        GenerarCierreDolares ("85")
    Else
        lArrMnt(0) = "CANCELACION_9"
        lArrMnt(1) = gsEmpresa
        lArrMnt(2) = gsAnio
        lArrMnt(3) = Me.tdbcLibro.BoundText
        lArrMnt(4) = 1
        lArrMnt(5) = gsUsuario
        lArrMnt(6) = "COM"
        lArrMnt(7) = "13" 'ya no es necesario por q en el SP se asigan valor
        lArrMnt(8) = gsByMoneda
        lArrMnt(9) = dtpFechaCierre.Value
    

        cmdGenerarCierre.Enabled = False
        cmdSalir(1).Enabled = False
        DoEvents
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, sp, lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            cmdGenerarCierre.Enabled = True
            cmdSalir(1).Enabled = True
            Exit Sub
        End If
   
        
        Set clsMante = Nothing
        
        DoEvents
        Call ActualizaSaldos
        DoEvents
         End If
        
        Screen.MousePointer = vbDefault
        Mensajes "Los asientos de cierre del ejercicio se crearon correctamente en el periodo de Ajuste", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
    cmdGenerarCierre.Enabled = True
    cmdSalir(1).Enabled = True
End Sub

Private Sub ActualizaSaldos()
        'Mensajes "SE INICIARA LA ACTUALIZACION DE CUENTAS DE DESTINO", vbOKOnly + vbExclamation
'        frmPrcActualizaDestino.Show
'        frmPrcActualizaDestino.cmdProcesar.Visible = False
'        DoEvents
'        frmPrcActualizaDestino.chkMes.Value = vbChecked
'        frmPrcActualizaDestino.chkMes.Enabled = False
'        frmPrcActualizaDestino.tdbcMes.BoundText = "14"
'        DoEvents
'        frmPrcActualizaDestino.gsMensaje = False
'        frmPrcActualizaDestino.gsSinSaldos = True
'        DoEvents
'        frmPrcActualizaDestino.Procesar
'        DoEvents
'        frmPrcActualizaDestino.Cerrar
'
'        DoEvents
        
        'Mensajes "SE INICIARA LA ACTUALIZACION DE SALDOS", vbOKOnly + vbExclamation
        frmPrcActualizaSaldos.Show
        frmPrcActualizaSaldos.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaSaldos.chkMes.Value = vbChecked
        frmPrcActualizaSaldos.chkMes.Enabled = False
        DoEvents
        frmPrcActualizaSaldos.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaSaldos.gsMensaje = False
        frmPrcActualizaSaldos.Procesar
        DoEvents
        frmPrcActualizaSaldos.Cerrar

End Sub

Private Function BuscaAsientosApertura() As Boolean
    Dim sql As String
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas

    
    sql = "select count(*) as registros from CNC_asiento_voucher " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & CE(NE(gsAnio) + 1) & "' and ase_cdeleted<>'*'"
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    If Not rsAddItem Is Nothing And rsAddItem.State = adStateOpen Then
        Do While Not rsAddItem.EOF
            If NE(rsAddItem!registros) > 0 Then
                BuscaAsientosApertura = True
            Else
                BuscaAsientosApertura = False
            End If
            rsAddItem.MoveNext
        Loop
    Else
        BuscaAsientosApertura = False
    End If
    
    If BuscaAsientosApertura = False Then
        Mensajes "Primero GENERE EL ASIENTO DE APERTURA del año siguiente, luego continue con este proceso", vbOKOnly + vbExclamation
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing


End Function


Private Sub cmdGenerarImpuesto_Click()



    If CE(tdbcLibro.Text) = "" Then
        Mensajes "Cree el libro de diferencia en cambio y configurelo en Parametros iniciales", vbOKOnly + vbInformation
        SSTCentroCosto.Tab = 0
        pSetFocus tdbcLibro
        Exit Sub

    End If


    If BuscaCuentasCierre = False Then Exit Sub
    
    Dim respuesta As String
    
    If NE(tdbHTRentaEj) = 0 And NE(tdbHTPerdidaEj) = 0 Then
        Mensajes "La perdida o la renta del ejercicio no puede ser CERO, calcule o ingrese los importes", vbOKOnly + vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbNormal
    
    respuesta = MsgBox("Desea Generar los asientos de cierre del Impuesto a la Renta", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Cierre")
    If respuesta = vbYes Then
    
        'If BuscaAsientosApertura = False Then
        '    Exit Sub
        'End If
    
        Screen.MousePointer = vbHourglass
        Dim lArrMnt(12) As Variant
        lArrMnt(0) = "'" & gsEmpresa & "'"
        lArrMnt(1) = "'" & gsAnio & "'"
        lArrMnt(2) = "'" & Me.tdbcLibro.BoundText & "'"
        lArrMnt(3) = NE(tdbHTtc)
        lArrMnt(4) = "'" & gsUsuario & "'"
        lArrMnt(5) = "'COM'"
        lArrMnt(6) = "'14'" 'ya no es necesario por q en el SP se asigan valor
        
        lArrMnt(7) = NE(tdbHTParticip)
        lArrMnt(8) = NE(tdbHTImpRenta)
        lArrMnt(9) = NE(tdbHTReservaLegal)
        lArrMnt(10) = NE(tdbHTUtilNetaEjer)
        lArrMnt(11) = NE(tdbHTUtilcont) - NE(tdbHTPerdCont) - NE(tdbHTParticip) - NE(tdbHTImpRenta) - NE(tdbHTReservaLegal)
        lArrMnt(12) = gsByMoneda
        
        Dim obj As New ClsFuncionesExecute
        obj.Mant_Tablas lArrMnt, "spCn_GeneraAsientoEjercicioRenta", 11
        Set obj = Nothing
        
        DoEvents
        ActualizaSaldos
        DoEvents
        
        Screen.MousePointer = vbDefault
        Mensajes "Los asientos de cierre de impuestos a la renta se crearon correctamente", vbInformation
    End If
End Sub

Private Sub cmdGrabarDatos_Click()
    GrabaHojaTrabajo
End Sub

Private Sub cmdImprimir_Click()
    Imprimir
End Sub

Private Sub cmdLeerDatos_Click()
 BuscaHojaTrabajo
End Sub


Private Sub ActivarControlPersup(ByRef Control As Object, Valor As Boolean)
    Control.Enabled = Valor
    
    If Valor = True Then
       Control.BackColor = &HC0FFFF
    Else
       Control.BackColor = gsColorDesactivado
    End If
End Sub


Private Sub cmdsalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Imprimir()
    Dim matriz(24) As Variant
    Dim Titulo As String
    Screen.MousePointer = vbHourglass
    Titulo = "HOJA DE TRABAJO DE DETERMINACION DE IMPUESTO A LA RENTA " & gsAnio
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo;" & Titulo & ";True"
    matriz(2) = "@Accion;BUSCARREGISTRO;True"
    matriz(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(4) = "@Pan_cAnio;" & gsAnio & ";True"

    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptHojaTrabajo.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    
    Dim sqlcombos As String
    Dim registros As Integer
    SSTCentroCosto.Tab = 0
    pCargaCfgLibro
    DoEvents

    Call Centrar_form(Me)
    
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND LIB_CTIPOLIBRO ='" & lsLibroCierre & "' AND Pan_cAnio = '" & gsAnio & "' ORDER BY LIB_CDESCRIPCION "
    registros = LlenarComboAddItem(tdbcLibro, sqlcombos)
    
    If registros <= 0 Then
        Mensajes "No se creo el libro de cierre o no se configuro en parametros iniciales, verifique los datos necesarios", vbOKOnly + vbInformation
    End If
    
    DoEvents
    
   ' If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
   '     Me.cmdGenerarCierre.Enabled = False
   '     Me.cmdGenerarImpuesto.Enabled = False
   ' Else
   '     Me.cmdGenerarCierre.Enabled = True
   '     Me.cmdGenerarImpuesto.Enabled = True
   ' End If
    
    
    Call cmdLeerDatos_Click
    DoEvents
    Call ControlAbs(tdbHTUtilcont)
    Call ControlAbs(tdbHTPerdCont)
    
    Call chkReserva_Click
    Call chkPorcImp_Click
    Call chkPorcParticip_Click
    
    On Error Resume Next
    dtpFechaCierre.Value = UltimoDiaMes("12", gsAnio)
    tdbcLibro.Bookmark = 0
    tdbcLibro.ReBind
    
End Sub

Private Sub CargaTabla()
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

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub SSTCentroCosto_Click(PreviousTab As Integer)
    If SSTCentroCosto.Tab = 1 Then
        cmdCargarDatos_Click
    End If
End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub


Private Sub CalculaRentayPerdidaEjercicio()
    On Error GoTo serror
    Dim Valor  As Double
    Valor = NE(tdbHTUtilcont) - NE(tdbHTPerdCont) + NE(tdbHTAdic) + NE(tdbHTDeducc)
    
    'tdbHTRentaEj = 0
    'tdbHTPerdidaEj = 0
    'tdbHTRentaNetaAntesImp = 0
    
    If Valor <= 0 Then
        tdbHTRentaEj = "0.00"
        tdbHTPerdidaEj = Valor
        ControlAbs tdbHTPerdidaEj
    Else
        tdbHTPerdidaEj = "0.00"
        tdbHTRentaEj = Valor
        ControlAbs tdbHTRentaEj
    End If


    Valor = NE(tdbHTRentaEj) - NE(tdbHTPerdidaEj) + NE(tdbHTPerdidaNeta)
    
    If NE(tdbHTPerdidaEj) >= 0 Then
        
        If tdbHTRentaNetaAntesImp.Enabled = False Then
            Valor = 0
        End If
    Else
        Valor = NE(tdbHTRentaEj) - NE(tdbHTPerdidaEj) + NE(tdbHTPerdidaNeta)
    End If
    
    
    
    If Valor >= 0 Then
       tdbHTRentaNetaAntesImp.Text = Abs(Valor)
       ControlAbs tdbHTRentaNetaAntesImp
       
       If Valor = 0 Then tdbHTRentaNetaAntesImp = "0.00"
    Else
        tdbHTRentaNetaAntesImp.Text = Valor
    End If
    Exit Sub
serror:
    Mensajes Err.Description, vbInformation + vbOKOnly
End Sub

Private Sub tdbHTAdic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ControlAbs tdbHTAdic
        CalculaRentayPerdidaEjercicio
        
        
        pSendKeys "{TAB}"
    End If
End Sub

Private Sub tdbHTAdic_LostFocus()
    If Not IsNumeric(tdbHTAdic) Then
       Mensajes "Ingrese un campo numerico", vbOKOnly + vbInformation
       tdbHTAdic = "0.00"
       Exit Sub
    End If

    ControlAbs tdbHTAdic
    CalculaRentayPerdidaEjercicio
End Sub


Private Sub tdbHTDeducc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdbHTDeducc = Abs(NE(tdbHTDeducc)) * -1
        tdbHTDeducc = Format(NE(tdbHTDeducc), "###,###,##0.00")
        CalculaRentayPerdidaEjercicio
        pSendKeys "{TAB}"
    End If

End Sub

Private Sub tdbHTDeducc_LostFocus()
    If Not IsNumeric(tdbHTDeducc) Then
       Mensajes "Ingrese un campo numerico", vbOKOnly + vbInformation
       tdbHTDeducc = "0.00"
       Exit Sub
    End If

    tdbHTDeducc = Abs(NE(tdbHTDeducc)) * -1
    tdbHTDeducc = Format(NE(tdbHTDeducc), "###,###,##0.00")
    CalculaRentayPerdidaEjercicio
End Sub

Private Sub tdbHTPerdCont_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSendKeys "{TAB}"
    End If

End Sub

Private Sub tdbHTPerdidaEj_Change()
    If NE(tdbHTPerdidaEj) > 0 Then
        ActivarControlPersup tdbHTPerdidaNeta, False
        ActivarControlPersup tdbHTRentaNetaAntesImp, False
    
        ActivarControlPersup tdbHTParticip, False
        ActivarControlPersup tdbHTRentaNetaImponible, False
        ActivarControlPersup tdbHTImpRenta, False
        ActivarControlPersup tdbHTUtilNetaEjer, False
        ActivarControlPersup tdbHTReservaLegal, False
        ActivarControlPersup tdbHTUtilDist, False
        
        tdbHTParticip.Text = "0.00"
        tdbHTRentaNetaImponible.Text = "0.00"
        tdbHTImpRenta.Text = "0.00"
        tdbHTUtilNetaEjer.Text = "0.00"
        tdbHTReservaLegal.Text = "0.00"
        tdbHTUtilDist.Text = "0.00"
        tdbHTRentaNetaAntesImp.Text = "0.00"
        tdbHTPerdidaNeta.Text = "0.00"
        
        chkPorcParticip.Value = vbUnchecked
        
    Else
        ActivarControlPersup tdbHTPerdidaNeta, True
        ActivarControlPersup tdbHTRentaNetaAntesImp, True
    
        ActivarControlPersup tdbHTParticip, True
        ActivarControlPersup tdbHTRentaNetaImponible, True
        ActivarControlPersup tdbHTImpRenta, True
        ActivarControlPersup tdbHTUtilNetaEjer, True
        ActivarControlPersup tdbHTReservaLegal, True
        ActivarControlPersup tdbHTUtilDist, True
    
        tdbHTPerdidaNeta.BackColor = gsColorActivado
    End If
    
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click
End Sub

Private Sub tdbHTPerdidaEj_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSendKeys "{TAB}"
    End If

End Sub

Private Sub tdbHTPerdidaNeta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Abs(NE(tdbHTPerdidaNeta)) > Abs(NE(tdbHTRentaEj)) Then
            Mensajes "La perdida neta compensable no puede ser mayor a la Utilidad del ejercicio", vbOKOnly + vbInformation
            tdbHTPerdidaNeta = "0.00"
            Exit Sub
        End If
        
        ControlAbs tdbHTPerdidaNeta
        If tdbHTPerdidaNeta <> 0 Then tdbHTPerdidaNeta = "-" & tdbHTPerdidaNeta
        CalculaRentayPerdidaEjercicio
        pSetFocus tdbHTRentaNetaAntesImp
        
        
        
    End If

    

End Sub

Private Sub tdbHTPerdidaNeta_LostFocus()
    If Not IsNumeric(tdbHTPerdidaNeta) Then
       Mensajes "Ingrese un campo numerico", vbOKOnly + vbInformation
       tdbHTPerdidaNeta = "0.00"
       Exit Sub
    End If


        If Abs(NE(tdbHTPerdidaNeta)) > Abs(NE(tdbHTRentaEj)) Then
            Mensajes "La perdida neta compensable no puede ser mayor a la Utilidad del ejercicio", vbOKOnly + vbInformation
            tdbHTPerdidaNeta = "0.00"
            Exit Sub
        End If

    ControlAbs tdbHTPerdidaNeta
    If NE(tdbHTPerdidaNeta) <> 0 Then tdbHTPerdidaNeta = "-" & tdbHTPerdidaNeta

    CalculaRentayPerdidaEjercicio
End Sub

Private Sub tdbHTRentaEj_Change()
    If NE(tdbHTRentaEj) <= 0 Then
        ActivarControlPersup tdbHTPerdidaNeta, False
        ActivarControlPersup tdbHTRentaNetaAntesImp, False
    
        ActivarControlPersup tdbHTParticip, False
        ActivarControlPersup tdbHTRentaNetaImponible, False
        ActivarControlPersup tdbHTImpRenta, False
        ActivarControlPersup tdbHTUtilNetaEjer, False
        ActivarControlPersup tdbHTReservaLegal, False
        ActivarControlPersup tdbHTUtilDist, False
        
        tdbHTParticip.Text = "0.00"
        tdbHTRentaNetaImponible.Text = "0.00"
        tdbHTImpRenta.Text = "0.00"
        tdbHTUtilNetaEjer.Text = "0.00"
        tdbHTReservaLegal.Text = "0.00"
        tdbHTUtilDist.Text = "0.00"
        tdbHTRentaNetaAntesImp.Text = "0.00"
        tdbHTPerdidaNeta.Text = "0.00"
        
        chkPorcParticip.Value = vbUnchecked
        
    Else
        ActivarControlPersup tdbHTPerdidaNeta, True
        ActivarControlPersup tdbHTRentaNetaAntesImp, True
    
        ActivarControlPersup tdbHTParticip, True
        ActivarControlPersup tdbHTRentaNetaImponible, True
        ActivarControlPersup tdbHTImpRenta, True
        ActivarControlPersup tdbHTUtilNetaEjer, True
        ActivarControlPersup tdbHTReservaLegal, True
        ActivarControlPersup tdbHTUtilDist, True
    
        tdbHTPerdidaNeta.BackColor = gsColorActivado
    End If
    
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click
    
End Sub

Private Sub tdbHTRentaEj_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ControlAbs tdbHTAdic
        pSendKeys "{TAB}"
    End If

End Sub

Private Sub tdbHTRentaNetaAntesImp_Change()
    If NE(tdbHTRentaNetaAntesImp) <= 0 Then
        ActivarControlPersup tdbHTParticip, False
        ActivarControlPersup tdbHTRentaNetaImponible, False
        ActivarControlPersup tdbHTImpRenta, False
        ActivarControlPersup tdbHTUtilNetaEjer, False
        ActivarControlPersup tdbHTReservaLegal, False
        ActivarControlPersup tdbHTUtilDist, False
        
        tdbHTParticip.Text = "0.00"
        tdbHTRentaNetaImponible.Text = "0.00"
        tdbHTImpRenta.Text = "0.00"
        tdbHTUtilNetaEjer.Text = "0.00"
        tdbHTReservaLegal.Text = "0.00"
        tdbHTUtilDist.Text = "0.00"
        
    Else
        ActivarControlPersup tdbHTParticip, True
        ActivarControlPersup tdbHTRentaNetaImponible, True
        ActivarControlPersup tdbHTImpRenta, True
        ActivarControlPersup tdbHTUtilNetaEjer, True
        ActivarControlPersup tdbHTReservaLegal, True
        ActivarControlPersup tdbHTUtilDist, True
    
    End If
    
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click

End Sub

Private Sub tdbHTRentaNetaAntesImp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSendKeys "{TAB}"
    End If

End Sub

Private Sub tdbHTUtilcont_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSendKeys "{TAB}"
    End If

    
End Sub

'Private Sub tdbnTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim TipoCambio As Double
'    If KeyCode = 13 Then
'        If NE(tdbnTipoCambio.Value) = 0 Then
'            Mensajes "Ingrese un tipo de cambio valido", vbOKOnly + vbInformation
'            psetfocus tdbnTipoCambio
'            KeyCode = 0
'        Else
'            TipoCambio = CargaTCxMeses
'            If TipoCambio = 0 Then
'
'                ActualizaTCMensual tdbcMes.BoundText
'            Else
'                If Me.tdbnTipoCambio.Value <> TipoCambio Then
'                    Mensajes "El tipo de cambio se actualizara al tipo de cambio ingresado para este mes", vbOKOnly + vbInformation
'                End If
'                Me.tdbnTipoCambio.Value = TipoCambio
'
'            End If
'        End If
'    End If
'
'End Sub




Private Sub tdbtPorcPartic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CalculaRentayPerdidaEjercicio
    End If
End Sub

Private Sub tdbtPorcPartic_LostFocus()
    If NE(tdbtPorcPartic) = 0 Then chkPorcParticip.Value = vbUnchecked
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click
    
End Sub


Private Sub tdbtPorcRenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CalculaRentayPerdidaEjercicio
    End If

End Sub

Private Sub tdbtPorcRenta_LostFocus()
    If NE(tdbtPorcRenta) = 0 Then chkPorcImp.Value = vbUnchecked
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click

End Sub


Private Sub tdbtPorcReserva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CalculaRentayPerdidaEjercicio
    End If

End Sub

Private Sub tdbtPorcReserva_LostFocus()
    If NE(tdbtPorcReserva) = 0 Then chkReserva.Value = vbUnchecked
    chkReserva_Click
    chkPorcImp_Click
    chkPorcParticip_Click

End Sub
