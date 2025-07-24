VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Begin VB.Form frmConsultarLE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Libros Electrónicos"
   ClientHeight    =   6588
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5724
   Icon            =   "frmConsultarLE.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6588
   ScaleWidth      =   5724
   Begin VB.Frame fratodo 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.ListBox LstLE 
         Height          =   4464
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar LE"
         Height          =   615
         Left            =   4200
         Picture         =   "frmConsultarLE.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Eliminar"
         Top             =   600
         Width           =   1095
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   2025
         TabIndex        =   3
         Top             =   840
         Width           =   1935
         _ExtentX        =   3408
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
         Splits(0).DividerStyle=   2
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=572"
         Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=677"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=614"
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
         EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   276.095
         AutoSize        =   -1  'True
         GapHeight       =   36.283
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
         _PropDict       =   $"frmConsultarLE.frx":1454
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
      Begin TrueOleDBList70.TDBCombo TDBLibro 
         Height          =   276
         Left            =   2028
         TabIndex        =   4
         Top             =   360
         Width           =   1932
         _ExtentX        =   3408
         _ExtentY        =   466
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
         Splits(0).DividerStyle=   2
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=572"
         Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=677"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=614"
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
         EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   276.095
         AutoSize        =   -1  'True
         GapHeight       =   36.283
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
         _PropDict       =   $"frmConsultarLE.frx":14DB
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=660"
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
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LIBRO:"
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
         Left            =   960
         TabIndex        =   6
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MES:"
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
         Left            =   960
         TabIndex        =   5
         Top             =   885
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmConsultarLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDir, strDireccionBckLE As String
Dim fechahoy As Date
Dim mes_fichero As String
Dim anio_fichero As String
Dim sArch As String
Dim mes_tmp As String
Dim clDatos As New ClsFuncionesExecute
Dim sql As String

Private Sub cmdEliminar_Click()
On Error GoTo Control
Dim rsAddItem As ADODB.Recordset
Dim fecha As Date

sql = "select fecCrea from CNT_lIBROSGENERADOS where Emp_cCodigo='" & gsEmpresa & "' and Pan_cAnio='" & gsAnio & "' and Per_cPeriodo='" & tdbcMes.BoundText & "' and Lib_cTipoLibro='" & TDBLibro.BoundText & "'"
Set rsAddItem = clDatos.fRetornaRS(sql)

If rsAddItem.RecordCount > 0 And Right(Left(LstLE.Text, 26), 5) <> "05030" Then
    If Not rsAddItem Is Nothing Then
        fecha = rsAddItem!fecCrea
    End If
Else
    If Right(Left(LstLE.Text, 26), 5) = "05030" Then
        If MsgBox("¿Desea eliminar el archivo seleccionado?", vbInformation + vbOKCancel, gsmodulo) = vbOK Then
            filedelete (sDir & "\" & LstLE.Text)
            filedelete (strDireccionBckLE & "\" & LstLE.Text)
            Call BuscarFileLE
        End If
    Else
        filedelete (sDir & "\" & LstLE.Text)
        filedelete (strDireccionBckLE & "\" & LstLE.Text)
        LstLE.Clear
        Call BuscarFileLE
    End If
    
    'Mensajes "No existe archivos generados para eliminar.", vbExclamation
    Set clDatos = Nothing
    Set rsAddItem = Nothing
    Exit Sub
End If

If (Year(fecha) = Year(fechahoy) And Month(fecha) = Month(fechahoy)) Then
    anio_fichero = Left(Right(LstLE.Text, 24), 4)
    If Len(LstLE.Text) > 37 Then anio_fichero = Left(Right(LstLE.Text, 26), 4)
    If gsRVIE = "1" And TDBLibro.BoundText = lsLibroVen Then
        If LstLE.ListCount > 1 And Len(LstLE.Text) > 37 Then
            anio_fichero = "SI"
        End If
    End If
    If anio_fichero <> "SI" Then
        If MsgBox("¿Desea eliminar y desbloquear el Periodo " & tdbcMes.Text & " del ejercicio " & anio_fichero & "?", vbInformation + vbOKCancel, gsmodulo) = vbOK Then
            anio_fichero = "SI"
            EliminaLE (tdbcMes.BoundText)
        End If
    End If
    If anio_fichero = "SI" Then
        filedelete (sDir & "\" & LstLE.Text)
        filedelete (strDireccionBckLE & "\" & LstLE.Text)
        
        Mensajes "El Libro Electrónico ha sido eliminado!", vbInformation
        LstLE.Clear
        Call BuscarFileLE
    End If
Else
    Mensajes "Solo puede eliminar Libros Electrónicos que se hayan generado hasta con un mes anterioridad." & vbCrLf & _
             "Modifique la Fecha de su Sistema Operativo al " & Format(fecha, "dd/mm/yyyy") & ", para eliminar el Libro", vbExclamation
End If

Set clDatos = Nothing
Set rsAddItem = Nothing
Exit Sub

Control:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Function EliminaLE(cPeriodo As String) As Boolean
    ' *** Generar el cierre
    Screen.MousePointer = vbHourglass
    On Local Error GoTo ErrorEjecucion

    
    If TDBLibro.BoundText <> "04" Or TDBLibro.BoundText <> "DS" Then
        sql = "delete from CNT_lIBROSGENERADOS where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and Lib_cTipolibro =  '" & TDBLibro.BoundText & "'"
        clDatos.pEjecutaSQL (sql)
    End If
    
    If TDBLibro.BoundText = "03" Then '  Elimina el Diario Detalle del Plan de Cuentas
        sql = "delete from CNT_lIBROSGENERADOS where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and Lib_cTipolibro =  'LD'"
        clDatos.pEjecutaSQL (sql)
    End If
    
    
    If TDBLibro.BoundText = "03" Then
        If TDBLibro.BoundText = "03" And cPeriodo <> "12" Then
            'Diario
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '" & TDBLibro.BoundText & "'"
            clDatos.pEjecutaSQL (sql)
    
            'Ingreso
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '02'"
            clDatos.pEjecutaSQL (sql)
    
            'Egresos
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '04'"
            clDatos.pEjecutaSQL (sql)
            
            'Dif. Cambio
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '07'"
            clDatos.pEjecutaSQL (sql)
               
            'Apertura -- Desbloquea Apertura si no existe ningun Diario
            If cPeriodo = "01" Then
                sql = "SELECT * FROM CNT_CIERRE WITH(READUNCOMMITTED) WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '" & TDBLibro.BoundText & "' and Per_cPeriodo= '01' "
                If ExisteDato(sql) = False Then
                    sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '00' and TipoLibro =  '01'"
                    clDatos.pEjecutaSQL (sql)
                End If
            End If
        End If
        
        If TDBLibro.BoundText = "03" And cPeriodo = "12" Then
            'Diario
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '" & TDBLibro.BoundText & "'"
            clDatos.pEjecutaSQL (sql)
    
            'Ingreso
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '02'"
            clDatos.pEjecutaSQL (sql)
    
            'Egresos
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '04'"
            clDatos.pEjecutaSQL (sql)
            
            'Dif. Cambio
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '07'"
            clDatos.pEjecutaSQL (sql)
    
            'Cierre en Ajuste
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '13' and TipoLibro =  '08'"
            clDatos.pEjecutaSQL (sql)
            
            'Cierre en cierre
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '14' and TipoLibro =  '08'"
            clDatos.pEjecutaSQL (sql)
        End If
    Else
            sql = "delete from CNT_CIERRE where  Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio =  '" & gsAnio & "' and Per_cPeriodo= '" & tdbcMes.BoundText & "' and TipoLibro =  '" & TDBLibro.BoundText & "'"
            clDatos.pEjecutaSQL (sql)
    End If

    Screen.MousePointer = vbDefault
    Exit Function
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Private Sub Form_Load()
   Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
   Call Centrar_form(Me)
   
   Call Llenarlibros
   Call LlenarPeriodo
   Call BuscarFileLE
   fechahoy = Date
End Sub

Private Sub filedelete(filename As String)
Dim filesystemobject As Object
Set filesystemobject = CreateObject("Scripting.filesystemobject")
filesystemobject.DeleteFile filename, True
End Sub

Private Sub LlenarPeriodo()
    Dim i As Integer
    Dim cadena As String
    For i = 0 To 11
        tdbcMes.AddItem Format(i + 1, "00") & ";" & UCase(MonthName(i + 1))
    Next
    tdbcMes.Bookmark = 0
    tdbcMes.ListField = "column1"
    tdbcMes.BoundColumn = "column0"
End Sub

Private Sub Llenarlibros()
'    TDBLibro.AddItem "01" & ";" & "APERTURA"
'    TDBLibro.AddItem "02" & ";" & "CAJA INGRESOS"
    If gsDiarioSimplificado = 0 Then
        TDBLibro.AddItem "03" & ";" & "DIARIO"
        TDBLibro.AddItem "04" & ";" & "MAYOR"
    End If
    TDBLibro.AddItem "05" & ";" & "VENTAS"
    TDBLibro.AddItem "06" & ";" & "COMPRAS"
    If gsDiarioSimplificado = 1 Then TDBLibro.AddItem "DS" & ";" & "DIARIO SIMPLIFICADO"
'    TDBLibro.AddItem "07" & ";" & "DIFERENCIA DE CAMBIO"
'    TDBLibro.AddItem "08" & ";" & "CIERRE"
    TDBLibro.Bookmark = 0
    TDBLibro.ListField = "column1"
    TDBLibro.BoundColumn = "column0"
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
Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub tdbcMes_ItemChange()
LstLE.Clear
Call BuscarFileLE
End Sub

Private Sub TDBLibro_ItemChange()
LstLE.Clear
Call BuscarFileLE
End Sub

Private Sub BuscarFileLE()
' sDir = App.Path     ' Directorio de la aplicación
If TDBLibro.BoundText = lsLibroCom Then
    strDireccionBckLE = App.Path & "\Backup_LE\" & gsRUC & "-" & gsEmpresa & "\Compras"
    sDir = App.Path & "\Libros_Electronicos\" & gsRUC & "-" & gsEmpresa & "\Compras"
ElseIf TDBLibro.BoundText = lsLibroVen Then
    sDir = App.Path & "\Libros_Electronicos\" & gsRUC & "-" & gsEmpresa & "\Ventas"
    strDireccionBckLE = App.Path & "\Backup_LE\" & gsRUC & "-" & gsEmpresa & "\Ventas"
ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then
    sDir = App.Path & "\Libros_Electronicos\" & gsRUC & "-" & gsEmpresa & "\Diario"
    strDireccionBckLE = App.Path & "\Backup_LE\" & gsRUC & "-" & gsEmpresa & "\Diario"
ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then 'Mayor
    sDir = App.Path & "\Libros_Electronicos\" & gsRUC & "-" & gsEmpresa & "\Mayor"
    strDireccionBckLE = App.Path & "\Backup_LE\" & gsRUC & "-" & gsEmpresa & "\Mayor"
ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then 'Diario Simplificado
    sDir = App.Path & "\Libros_Electronicos\" & gsRUC & "-" & gsEmpresa & "\Diario_Simplificado"
    strDireccionBckLE = App.Path & "\Backup_LE\" & gsRUC & "-" & gsEmpresa & "\Diario_Simplificado"
Else
End If

Dim Posi As Integer
sArch = Dir(sDir & "\*.*")
mes_fichero = Left(Right(sArch, Posi), 2)
''hlp20231121
Do While sArch <> ""
'    nombre_fichero = Path.GetFileName(sArch)
    If Len(sArch) > 37 Then Posi = 22 Else Posi = 20
    If InStr(1, sArch, "CP") <> 0 Then
       Posi = 8
    End If
    If Len(sArch) = 38 Then
       Posi = 21
    End If
    mes_fichero = Left(Right(sArch, Posi), 2)
    anio_fichero = Left(Right(sArch, Posi + 4), 4)
    If mes_fichero = tdbcMes.BoundText And anio_fichero = gsAnio Then
        LstLE.AddItem sArch
    End If
    sArch = Dir
Loop
If LstLE.ListCount >= 1 Then
    LstLE.ListIndex = 0
End If
End Sub
