VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBusCoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Documentos PDB"
   ClientHeight    =   6525
   ClientLeft      =   1860
   ClientTop       =   3345
   ClientWidth     =   11505
   Icon            =   "frmBusCoa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11505
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   225
      TabIndex        =   8
      Top             =   195
      Width           =   7950
      Begin TDBText6Ctl.TDBText tdbtNombreEntidad 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   270
         Width           =   4620
         _Version        =   65536
         _ExtentX        =   8149
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":0ECA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":0F36
         Key             =   "frmBusCoa.frx":0F54
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
         Format          =   "A"
         FormatMode      =   0
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
      Begin TDBText6Ctl.TDBText tdbtCodigo 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   270
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":0F96
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":1002
         Key             =   "frmBusCoa.frx":1020
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
      Begin TDBText6Ctl.TDBText tdbtCuenta 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   990
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":1062
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":10CE
         Key             =   "frmBusCoa.frx":10EC
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
      Begin TDBText6Ctl.TDBText tdbtSerie 
         Height          =   315
         Left            =   3825
         TabIndex        =   3
         Top             =   990
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":112E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":119A
         Key             =   "frmBusCoa.frx":11B8
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
         Format          =   "A9"
         FormatMode      =   0
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
      Begin TDBText6Ctl.TDBText tdbtNumero 
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   990
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":11FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":1266
         Key             =   "frmBusCoa.frx":1284
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
         Format          =   "A9"
         FormatMode      =   0
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
      Begin TDBText6Ctl.TDBText tdbtRuc 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   630
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "frmBusCoa.frx":12B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusCoa.frx":1324
         Key             =   "frmBusCoa.frx":1342
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   14
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label4 
         Caption         =   "Ruc"
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
         Left            =   225
         TabIndex        =   12
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
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
         Left            =   4860
         TabIndex        =   11
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta"
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
         Left            =   225
         TabIndex        =   10
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
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
         Left            =   225
         TabIndex        =   9
         Top             =   315
         Width           =   1095
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   330
      Left            =   2925
      TabIndex        =   7
      Top             =   3645
      Visible         =   0   'False
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   582
      Calculator      =   "frmBusCoa.frx":1384
      Caption         =   "frmBusCoa.frx":13A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBusCoa.frx":1410
      Keys            =   "frmBusCoa.frx":142E
      Spin            =   "frmBusCoa.frx":1478
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,##0.00"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999999
      MinValue        =   -10000
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1802698757
      MinValueVT      =   1769209861
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgProvisiones 
      Height          =   4170
      Left            =   180
      TabIndex        =   6
      Top             =   1770
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   7355
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Id"
      Columns(0).DataField=   "Ase_cNummov"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Empresa"
      Columns(1).DataField=   "Emp_cCodigo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Año"
      Columns(2).DataField=   "Pan_cAnio"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Periodo"
      Columns(3).DataField=   "Per_cPeriodo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Libro"
      Columns(4).DataField=   "Lib_cTipoLibro"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Voucher"
      Columns(5).DataField=   "Ase_nVoucher"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Cuenta"
      Columns(6).DataField=   "Pla_cCuentaContable"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Tipo"
      Columns(7).DataField=   "Ten_cTipoEntidad"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Codigo"
      Columns(8).DataField=   "Ent_cCodEntidad"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Ruc"
      Columns(9).DataField=   "Ent_nRuc"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Razon Social"
      Columns(10).DataField=   "Ent_cPersona"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "TD"
      Columns(11).DataField=   "Asd_cTipoDoc"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Serie"
      Columns(12).DataField=   "Asd_cSerieDoc"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Numero"
      Columns(13).DataField=   "Asd_cNumDoc"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Total Soles"
      Columns(14).DataField=   "Soles"
      Columns(14).NumberFormat=   "External Editor"
      Columns(14).ExternalEditor=   "TDBNumber1"
      Columns(14).ExternalEditor.vt=   8
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "TC"
      Columns(15).DataField=   "Asd_nTipoCambio"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Total MonExt"
      Columns(16).DataField=   "Dolares"
      Columns(16).NumberFormat=   "External Editor"
      Columns(16).ExternalEditor=   "TDBNumber1"
      Columns(16).ExternalEditor.vt=   8
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Fecha"
      Columns(17).DataField=   "Asd_dFecDoc"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Tipo"
      Columns(18).DataField=   "Com_ctipoIgv"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Glosa"
      Columns(19).DataField=   "asd_cGlosa"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "NumMov"
      Columns(20).DataField=   "Ase_cNummov"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Item"
      Columns(21).DataField=   "asd_nitem"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Moneda"
      Columns(22).DataField=   "asd_ctipomoneda"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Retencion"
      Columns(23).DataField=   "asd_cretencion"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "fecha spot"
      Columns(24).DataField=   "asd_dfechaspot"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "num spot"
      Columns(25).DataField=   "asd_cnumspot"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "asd_ctipodocref"
      Columns(26).DataField=   "asd_ctipodocref"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "asd_cseriedocref"
      Columns(27).DataField=   "asd_cseriedocref"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "asd_cnumdocref"
      Columns(28).DataField=   "asd_cnumdocref"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "asd_dfecdocref"
      Columns(29).DataField=   "asd_dfecdocref"
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "asd_cformapago"
      Columns(30).DataField=   "asd_cformapago"
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "Ent_cFlagPersona"
      Columns(31).DataField=   "Ent_cFlagPersona"
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "Ent_cTipoDoc"
      Columns(32).DataField=   "Ent_cTipoDoc"
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "Asd_cBaseImp"
      Columns(33).DataField=   "Asd_cBaseImp"
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "Ent_cApaterno"
      Columns(34).DataField=   "Ent_cApaterno"
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "Ent_cAmaterno"
      Columns(35).DataField=   "Ent_cAmaterno"
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "Ent_cNombres"
      Columns(36).DataField=   "Ent_cNombres"
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   37
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=37"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=661"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=582"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=926"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=847"
      Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=2143"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2064"
      Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(42)=   "Column(6).Width=2990"
      Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2910"
      Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(48)=   "Column(7).Width=661"
      Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=582"
      Splits(0)._ColumnProps(51)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(54)=   "Column(8).Width=1244"
      Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=1164"
      Splits(0)._ColumnProps(57)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(59)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(60)=   "Column(9).Width=3545"
      Splits(0)._ColumnProps(61)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(9)._WidthInPix=3466"
      Splits(0)._ColumnProps(63)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(64)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(65)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(66)=   "Column(10).Width=5133"
      Splits(0)._ColumnProps(67)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(10)._WidthInPix=5054"
      Splits(0)._ColumnProps(69)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(70)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(71)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(72)=   "Column(11).Width=582"
      Splits(0)._ColumnProps(73)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(11)._WidthInPix=503"
      Splits(0)._ColumnProps(75)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(76)=   "Column(11)._ColStyle=513"
      Splits(0)._ColumnProps(77)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(78)=   "Column(12).Width=1402"
      Splits(0)._ColumnProps(79)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(12)._WidthInPix=1323"
      Splits(0)._ColumnProps(81)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(82)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(83)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(84)=   "Column(13).Width=2117"
      Splits(0)._ColumnProps(85)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(13)._WidthInPix=2037"
      Splits(0)._ColumnProps(87)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(88)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(89)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(90)=   "Column(14).Width=1588"
      Splits(0)._ColumnProps(91)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(14)._WidthInPix=1508"
      Splits(0)._ColumnProps(93)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(94)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(95)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(96)=   "Column(15).Width=794"
      Splits(0)._ColumnProps(97)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(15)._WidthInPix=714"
      Splits(0)._ColumnProps(99)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(100)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(102)=   "Column(16).Width=1852"
      Splits(0)._ColumnProps(103)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(104)=   "Column(16)._WidthInPix=1773"
      Splits(0)._ColumnProps(105)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(106)=   "Column(16)._ColStyle=514"
      Splits(0)._ColumnProps(107)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(108)=   "Column(17).Width=1693"
      Splits(0)._ColumnProps(109)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(17)._WidthInPix=1614"
      Splits(0)._ColumnProps(111)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(112)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(113)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(114)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(115)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(117)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(118)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(119)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(120)=   "Column(18).AllowFocus=0"
      Splits(0)._ColumnProps(121)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(122)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(123)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(125)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(126)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(127)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(128)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(129)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(130)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(131)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(132)=   "Column(20).AllowSizing=0"
      Splits(0)._ColumnProps(133)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(134)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(135)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(136)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(137)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(138)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(140)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(141)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(142)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(143)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(144)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(145)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(146)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(147)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(148)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(149)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(150)=   "Column(22).AllowSizing=0"
      Splits(0)._ColumnProps(151)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(152)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(153)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(154)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(155)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(156)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(158)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(159)=   "Column(23).AllowSizing=0"
      Splits(0)._ColumnProps(160)=   "Column(23)._ColStyle=516"
      Splits(0)._ColumnProps(161)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(162)=   "Column(23).AllowFocus=0"
      Splits(0)._ColumnProps(163)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(164)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(165)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(167)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(168)=   "Column(24).AllowSizing=0"
      Splits(0)._ColumnProps(169)=   "Column(24)._ColStyle=516"
      Splits(0)._ColumnProps(170)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(171)=   "Column(24).AllowFocus=0"
      Splits(0)._ColumnProps(172)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(173)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(174)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(176)=   "Column(25)._EditAlways=0"
      Splits(0)._ColumnProps(177)=   "Column(25).AllowSizing=0"
      Splits(0)._ColumnProps(178)=   "Column(25)._ColStyle=516"
      Splits(0)._ColumnProps(179)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(180)=   "Column(25).AllowFocus=0"
      Splits(0)._ColumnProps(181)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(182)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(183)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(184)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(185)=   "Column(26)._EditAlways=0"
      Splits(0)._ColumnProps(186)=   "Column(26).AllowSizing=0"
      Splits(0)._ColumnProps(187)=   "Column(26)._ColStyle=516"
      Splits(0)._ColumnProps(188)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(189)=   "Column(26).AllowFocus=0"
      Splits(0)._ColumnProps(190)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(191)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(192)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(194)=   "Column(27)._EditAlways=0"
      Splits(0)._ColumnProps(195)=   "Column(27).AllowSizing=0"
      Splits(0)._ColumnProps(196)=   "Column(27)._ColStyle=516"
      Splits(0)._ColumnProps(197)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(198)=   "Column(27).AllowFocus=0"
      Splits(0)._ColumnProps(199)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(200)=   "Column(28).Width=2725"
      Splits(0)._ColumnProps(201)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(202)=   "Column(28)._WidthInPix=2646"
      Splits(0)._ColumnProps(203)=   "Column(28)._EditAlways=0"
      Splits(0)._ColumnProps(204)=   "Column(28).AllowSizing=0"
      Splits(0)._ColumnProps(205)=   "Column(28)._ColStyle=516"
      Splits(0)._ColumnProps(206)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(207)=   "Column(28).AllowFocus=0"
      Splits(0)._ColumnProps(208)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(209)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(210)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(212)=   "Column(29)._EditAlways=0"
      Splits(0)._ColumnProps(213)=   "Column(29).AllowSizing=0"
      Splits(0)._ColumnProps(214)=   "Column(29)._ColStyle=516"
      Splits(0)._ColumnProps(215)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(216)=   "Column(29).AllowFocus=0"
      Splits(0)._ColumnProps(217)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(218)=   "Column(30).Width=2725"
      Splits(0)._ColumnProps(219)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(220)=   "Column(30)._WidthInPix=2646"
      Splits(0)._ColumnProps(221)=   "Column(30)._EditAlways=0"
      Splits(0)._ColumnProps(222)=   "Column(30).AllowSizing=0"
      Splits(0)._ColumnProps(223)=   "Column(30)._ColStyle=516"
      Splits(0)._ColumnProps(224)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(225)=   "Column(30).AllowFocus=0"
      Splits(0)._ColumnProps(226)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(227)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(228)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(229)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(230)=   "Column(31)._EditAlways=0"
      Splits(0)._ColumnProps(231)=   "Column(31).AllowSizing=0"
      Splits(0)._ColumnProps(232)=   "Column(31)._ColStyle=516"
      Splits(0)._ColumnProps(233)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(234)=   "Column(31).AllowFocus=0"
      Splits(0)._ColumnProps(235)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(236)=   "Column(32).Width=2725"
      Splits(0)._ColumnProps(237)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(238)=   "Column(32)._WidthInPix=2646"
      Splits(0)._ColumnProps(239)=   "Column(32)._EditAlways=0"
      Splits(0)._ColumnProps(240)=   "Column(32).AllowSizing=0"
      Splits(0)._ColumnProps(241)=   "Column(32)._ColStyle=516"
      Splits(0)._ColumnProps(242)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(243)=   "Column(32).AllowFocus=0"
      Splits(0)._ColumnProps(244)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(245)=   "Column(33).Width=2725"
      Splits(0)._ColumnProps(246)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(33)._WidthInPix=2646"
      Splits(0)._ColumnProps(248)=   "Column(33)._EditAlways=0"
      Splits(0)._ColumnProps(249)=   "Column(33).AllowSizing=0"
      Splits(0)._ColumnProps(250)=   "Column(33)._ColStyle=516"
      Splits(0)._ColumnProps(251)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(252)=   "Column(33).AllowFocus=0"
      Splits(0)._ColumnProps(253)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(254)=   "Column(34).Width=2725"
      Splits(0)._ColumnProps(255)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(256)=   "Column(34)._WidthInPix=2646"
      Splits(0)._ColumnProps(257)=   "Column(34)._EditAlways=0"
      Splits(0)._ColumnProps(258)=   "Column(34)._ColStyle=516"
      Splits(0)._ColumnProps(259)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(260)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(261)=   "Column(35).Width=2725"
      Splits(0)._ColumnProps(262)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(35)._WidthInPix=2646"
      Splits(0)._ColumnProps(264)=   "Column(35)._EditAlways=0"
      Splits(0)._ColumnProps(265)=   "Column(35)._ColStyle=516"
      Splits(0)._ColumnProps(266)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(267)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(268)=   "Column(36).Width=2725"
      Splits(0)._ColumnProps(269)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(270)=   "Column(36)._WidthInPix=2646"
      Splits(0)._ColumnProps(271)=   "Column(36)._EditAlways=0"
      Splits(0)._ColumnProps(272)=   "Column(36)._ColStyle=516"
      Splits(0)._ColumnProps(273)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(274)=   "Column(36).Order=37"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=16,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H80000007&"
      _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000E&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Arial"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=Arial"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=122,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=119,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=120,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=121,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=118,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=138,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=135,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=136,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=137,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=134,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=131,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=132,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=133,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=130,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=126,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=123,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=124,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=125,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=50,.parent=13"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13,.alignment=1"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=32,.parent=13"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=29,.parent=14"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=30,.parent=15"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=31,.parent=17"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=98,.parent=13,.alignment=1"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=82,.parent=13"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=79,.parent=14"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=80,.parent=15"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=81,.parent=17"
      _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
      _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
      _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
      _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
      _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=28,.parent=13"
      _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=25,.parent=14"
      _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=26,.parent=15"
      _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=27,.parent=17"
      _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=46,.parent=13"
      _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=43,.parent=14"
      _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=44,.parent=15"
      _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=45,.parent=17"
      _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=54,.parent=13"
      _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=51,.parent=14"
      _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=52,.parent=15"
      _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=53,.parent=17"
      _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=74,.parent=13"
      _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=71,.parent=14"
      _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=72,.parent=15"
      _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=73,.parent=17"
      _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=102,.parent=13"
      _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=99,.parent=14"
      _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=100,.parent=15"
      _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=101,.parent=17"
      _StyleDefs(133) =   "Splits(0).Columns(24).Style:id=106,.parent=13"
      _StyleDefs(134) =   "Splits(0).Columns(24).HeadingStyle:id=103,.parent=14"
      _StyleDefs(135) =   "Splits(0).Columns(24).FooterStyle:id=104,.parent=15"
      _StyleDefs(136) =   "Splits(0).Columns(24).EditorStyle:id=105,.parent=17"
      _StyleDefs(137) =   "Splits(0).Columns(25).Style:id=114,.parent=13"
      _StyleDefs(138) =   "Splits(0).Columns(25).HeadingStyle:id=111,.parent=14"
      _StyleDefs(139) =   "Splits(0).Columns(25).FooterStyle:id=112,.parent=15"
      _StyleDefs(140) =   "Splits(0).Columns(25).EditorStyle:id=113,.parent=17"
      _StyleDefs(141) =   "Splits(0).Columns(26).Style:id=142,.parent=13"
      _StyleDefs(142) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=14"
      _StyleDefs(143) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=15"
      _StyleDefs(144) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=17"
      _StyleDefs(145) =   "Splits(0).Columns(27).Style:id=146,.parent=13"
      _StyleDefs(146) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
      _StyleDefs(147) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
      _StyleDefs(148) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
      _StyleDefs(149) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
      _StyleDefs(150) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
      _StyleDefs(151) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
      _StyleDefs(152) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
      _StyleDefs(153) =   "Splits(0).Columns(29).Style:id=154,.parent=13"
      _StyleDefs(154) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=14"
      _StyleDefs(155) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=15"
      _StyleDefs(156) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=17"
      _StyleDefs(157) =   "Splits(0).Columns(30).Style:id=158,.parent=13"
      _StyleDefs(158) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=14"
      _StyleDefs(159) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=15"
      _StyleDefs(160) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=17"
      _StyleDefs(161) =   "Splits(0).Columns(31).Style:id=162,.parent=13"
      _StyleDefs(162) =   "Splits(0).Columns(31).HeadingStyle:id=159,.parent=14"
      _StyleDefs(163) =   "Splits(0).Columns(31).FooterStyle:id=160,.parent=15"
      _StyleDefs(164) =   "Splits(0).Columns(31).EditorStyle:id=161,.parent=17"
      _StyleDefs(165) =   "Splits(0).Columns(32).Style:id=166,.parent=13"
      _StyleDefs(166) =   "Splits(0).Columns(32).HeadingStyle:id=163,.parent=14"
      _StyleDefs(167) =   "Splits(0).Columns(32).FooterStyle:id=164,.parent=15"
      _StyleDefs(168) =   "Splits(0).Columns(32).EditorStyle:id=165,.parent=17"
      _StyleDefs(169) =   "Splits(0).Columns(33).Style:id=170,.parent=13"
      _StyleDefs(170) =   "Splits(0).Columns(33).HeadingStyle:id=167,.parent=14"
      _StyleDefs(171) =   "Splits(0).Columns(33).FooterStyle:id=168,.parent=15"
      _StyleDefs(172) =   "Splits(0).Columns(33).EditorStyle:id=169,.parent=17"
      _StyleDefs(173) =   "Splits(0).Columns(34).Style:id=174,.parent=13"
      _StyleDefs(174) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=14"
      _StyleDefs(175) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=15"
      _StyleDefs(176) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=17"
      _StyleDefs(177) =   "Splits(0).Columns(35).Style:id=178,.parent=13"
      _StyleDefs(178) =   "Splits(0).Columns(35).HeadingStyle:id=175,.parent=14"
      _StyleDefs(179) =   "Splits(0).Columns(35).FooterStyle:id=176,.parent=15"
      _StyleDefs(180) =   "Splits(0).Columns(35).EditorStyle:id=177,.parent=17"
      _StyleDefs(181) =   "Splits(0).Columns(36).Style:id=182,.parent=13"
      _StyleDefs(182) =   "Splits(0).Columns(36).HeadingStyle:id=179,.parent=14"
      _StyleDefs(183) =   "Splits(0).Columns(36).FooterStyle:id=180,.parent=15"
      _StyleDefs(184) =   "Splits(0).Columns(36).EditorStyle:id=181,.parent=17"
      _StyleDefs(185) =   "Named:id=33:Normal"
      _StyleDefs(186) =   ":id=33,.parent=0"
      _StyleDefs(187) =   "Named:id=34:Heading"
      _StyleDefs(188) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(189) =   ":id=34,.wraptext=-1"
      _StyleDefs(190) =   "Named:id=35:Footing"
      _StyleDefs(191) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(192) =   "Named:id=36:Selected"
      _StyleDefs(193) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(194) =   "Named:id=37:Caption"
      _StyleDefs(195) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(196) =   "Named:id=38:HighlightRow"
      _StyleDefs(197) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(198) =   "Named:id=39:EvenRow"
      _StyleDefs(199) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(200) =   "Named:id=40:OddRow"
      _StyleDefs(201) =   ":id=40,.parent=33"
      _StyleDefs(202) =   "Named:id=41:RecordSelector"
      _StyleDefs(203) =   ":id=41,.parent=34"
      _StyleDefs(204) =   "Named:id=42:FilterBar"
      _StyleDefs(205) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione un documento haciendo doble click a un elemento de la lista"
      Height          =   825
      Left            =   8865
      TabIndex        =   13
      Top             =   735
      Width           =   2295
   End
End
Attribute VB_Name = "frmBusCoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmBusCoa
'    Project    : Contabilidad
'
'    Description: Formulario utilizado para la busqueda de movimientos para el PDB
'--------------------------------------------------------------------------------
Option Explicit
Dim lrsProvision As ADODB.Recordset
Public frmOrigen As Form
Public tabla As String
Public auxiliar As String
Public enUso As Boolean
Public Libro As String
Public pPeriodo As String
Public nDigitos As Integer

Public NombreOrigen As String
Public NombreBuscador As String


Dim gsTipoCOA As String

Dim pCuenta As String
Dim pEntidad As String
Dim pSerie As String
Dim pNumero As String

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Cuenta
' Description:       Propiedad de asignacion de cuenta
'
' Parameters :       vCuenta (String)
'--------------------------------------------------------------------------------
Public Property Let Cuenta(ByVal vCuenta As String)
     pCuenta = vCuenta
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Entidad
' Description:       Propiedad de asignacion de Entidad
'
' Parameters :       vEntidad (String)
'--------------------------------------------------------------------------------
Public Property Let Entidad(ByVal vEntidad As String)
     pEntidad = vEntidad
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Serie
' Description:       Propiedad de asignaciond e serie de documento
'
' Parameters :       vSerie (String)
'--------------------------------------------------------------------------------
Public Property Let Serie(ByVal vSerie As String)
     pSerie = vSerie
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       numero
' Description:       Propiedad de asignacion denumero de documento
'
' Parameters :       vNumero (String)
'--------------------------------------------------------------------------------
Public Property Let Numero(ByVal vNumero As String)
     pNumero = vNumero
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       pTipoCOA
' Description:       Propiedad de asignacion de tipo de PDB, si es de compras, ventas, caja.
'
' Parameters :       TipoCOA (String)
'--------------------------------------------------------------------------------
Public Property Let pTipoCOA(ByVal TipoCOA As String)
     gsTipoCOA = TipoCOA
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el formulario
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    KeyAscii = 0
    '    EnviaCodigo
    'End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    ' *** Inicializando datos y llenando provision
    Dim codMes As String
    
    Call Centrar_form(Me)
    
    codMes = "01"
    Call LlenaProvision
    'pSetFocus tdbtNombreEntidad
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaProvision
' Description:       Procedimiento de carga de provisiones que van al PDB
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaProvision()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Set lrsProvision = New ADODB.Recordset
    Set tdbgProvisiones.DataSource = Nothing
    
    sqlSp = "spCn_GrabaAsientoPDB 'BUSCA_PROVISIONES', '','','" & gsEmpresa & "', '" & gsAnio & "', '" & pPeriodo & "', '" & Libro & "', '" & auxiliar & "' "
    arrDatos = Array(sqlSp)
    Set lrsProvision = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsProvision Is Nothing Then
        ' *** Llenar grilla con el RecordSet
        tdbgProvisiones.DataSource = lrsProvision
        Exit Sub
    End If
    tdbgProvisiones.ReBind
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Resize
' Description:       Evento que se ejecuta almaximizar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error GoTo serror
        tdbgProvisiones.Width = Me.Width - 280
        tdbgProvisiones.Height = Me.Height - tdbgProvisiones.Top - 380
    End If
    
    Exit Sub
    
serror:


End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsProvision)
    frmOrigen.Enabled = True
    

    'Set frmOrigen = Nothing
    Set frmBusCoa = Nothing
    enUso = False
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tbdtCuenta_Change
' Description:       Evento que se ejecuta al cambiar el filtro del campo de cuenta
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tbdtCuenta_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgProvisiones_DblClick
' Description:       Evento que se ejecuta al hacer doble clic en la grilla de provisiones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbgProvisiones_DblClick()
    Call EnviaCodigo
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodigo_Change
' Description:       Evento que se ejecuta al cambiar el filtro de codigo
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCodigo_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodigo_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el filtro de codigo
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        Siguiente
        KeyCode = 0
    End If
    If KeyCode = 38 Then
        Anterior
        KeyCode = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCuenta_Change
' Description:       Evento que se ejecuta al cambiar el filtro de cuenta contable
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCuenta_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCuenta_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el filtro de cuenta contable
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        Siguiente
        KeyCode = 0
    End If
    If KeyCode = 38 Then
        Anterior
        KeyCode = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombreEntidad_Change
' Description:       Evento que se ejecuta al cambiar el filtro de descripcion de entidad
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNombreEntidad_Change()
    If gsKey = 219 Then
        tdbtNombreEntidad = Replace(tdbtNombreEntidad, "'", "")
        tdbtNombreEntidad.SelStart = Len(tdbtNombreEntidad)
    End If
    
    Call FiltrarRecordSet
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Siguiente
' Description:       Procedimiento que avanza el recordset de provisiones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Siguiente()
    tdbgProvisiones.MoveNext
    If tdbgProvisiones.EOF Then tdbgProvisiones.MoveLast
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Anterior
' Description:       Procedimiento que retrocede el recordset de provisiones
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Anterior()
    tdbgProvisiones.MovePrevious
    If tdbgProvisiones.BOF Then tdbgProvisiones.MoveFirst
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNombreEntidad_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el filtro de nombre de entidad
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNombreEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If KeyCode = 40 Then
        Siguiente
        KeyCode = 0
    End If
    If KeyCode = 38 Then
        Anterior
        KeyCode = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumero_Change
' Description:       Evento que se ejecuta al cambiar el filtro del numero de documento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNumero_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumero_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el filtro del numero de documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        Siguiente
        KeyCode = 0
    End If
    If KeyCode = 38 Then
        Anterior
        KeyCode = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtRuc_Change
' Description:       Evento que se ejecuta al cambiar el numero de ruc
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtRuc_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtSerie_Change
' Description:       Evento que se ejecuta al cambiar el filtro de la serie del documento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtSerie_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       FiltrarRecordSet
' Description:       Procedimiento que filtra el recordset segun los filtros digitados
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(5) As String
    Dim i As Integer
    cadena = ""
    If CE(tdbtCodigo) <> "" Then filtros(0) = "Ent_cCodEntidad like '*" & tdbtCodigo & "*'"
    If CE(tdbtNombreEntidad) <> "" Then filtros(1) = "Ent_cPersona like '*" & tdbtNombreEntidad & "*'"
    If CE(tdbtCuenta) <> "" Then filtros(2) = "Pla_cCuentaContable like '*" & tdbtCuenta & "*'"
    If CE(tdbtSerie) <> "" Then filtros(3) = "Asd_cSerieDoc like '*" & tdbtSerie & "*'"
    If CE(tdbtNumero) <> "" Then filtros(4) = "Asd_cNumDoc like '*" & tdbtNumero & "*'"
    If CE(tdbtRuc) <> "" Then filtros(5) = "ent_nRuc like '" & tdbtRuc & "*'"
    For i = 0 To 5
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    If lrsProvision Is Nothing = False Then lrsProvision.Filter = 0
    ' *** Filtrando segun campos
    If Not lrsProvision Is Nothing Then
        If Not (lrsProvision.BOF And lrsProvision.EOF) Then
            If CE(cadena) <> "" Then
                lrsProvision.Filter = cadena
            Else
                lrsProvision.Filter = 0
            End If
        End If
    End If
    

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       EnviaCodigo
' Description:       Ejecuta el procedimientode Recibir datos del formulario que lo invoca
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub EnviaCodigo()
    If Not lrsProvision Is Nothing Then
       If lrsProvision.RecordCount > 0 Then
          frmOrigen.Enabled = True
          frmOrigen.RecibirDatos "Provisiones", "", "", ""
       Else
          MsgBox "Código no existe, digite correctamente... "
       End If
    Else
        Unload Me
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtSerie_KeyDown
' Description:       Evento que se ejecuta al presionar  una tecla en el filtro de serie de documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        Siguiente
        KeyCode = 0
    End If
    If KeyCode = 38 Then
        Anterior
        KeyCode = 0
    End If
End Sub
