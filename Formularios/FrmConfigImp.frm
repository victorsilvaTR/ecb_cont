VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmConfigImp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Configurar Impresora :::"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "FrmConfigImp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Numero de Copias"
      Height          =   825
      Left            =   2880
      TabIndex        =   9
      Top             =   900
      Width           =   1680
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   315
         TabIndex        =   10
         Top             =   315
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccione Papel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   135
      TabIndex        =   6
      Top             =   1755
      Width           =   4455
      Begin VB.ComboBox CbPapel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   330
         Width           =   2970
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   8
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   135
      TabIndex        =   3
      Top             =   900
      Width           =   2625
      Begin VB.OptionButton Option2 
         Caption         =   "Vertical"
         Height          =   330
         Left            =   1545
         TabIndex        =   5
         Top             =   315
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Horizontal"
         Height          =   330
         Left            =   270
         TabIndex        =   4
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5685
      Begin VB.ComboBox CbImp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   435
      Left            =   4650
      TabIndex        =   12
      Top             =   2115
      Width           =   1170
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2064;767"
      Picture         =   "FrmConfigImp.frx":1982
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminar 
      Height          =   435
      Left            =   4650
      TabIndex        =   11
      Top             =   1620
      Width           =   1170
      Caption         =   " Aceptar"
      PicturePosition =   327683
      Size            =   "2064;767"
      Picture         =   "FrmConfigImp.frx":1F1C
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmConfigImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For Each x In Printers
    CbImp.AddItem x.DeviceName
    CbPapel.AddItem x.PaperSize
Next
vbPRPSLetter

End Sub
