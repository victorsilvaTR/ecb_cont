VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de ..."
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   Icon            =   "frmBusqueda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid70.TDBGrid tdbgListado 
      Height          =   4665
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   8229
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Producto"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Familia"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Clase"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Grupo"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tipo"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   4
      Splits(0).RecordSelectorWidth=   450
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=5741"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5662"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3519"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3440"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2037"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(28)=   "Column(4).Width=185"
      Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=106"
      Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=185"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=106"
      Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "BUSQUEDA"
      TabAction       =   2
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   2
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=160,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H8000000F&"
      _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000017&,.transparentBmp=0,.fgpicPosition=2,.bgpicMode=2"
      _StyleDefs(11)  =   ":id=4,.appearance=0,.ellipsis=0"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H80000002&"
      _StyleDefs(13)  =   ":id=2,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.bgcolor=&H80000005&"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.wraptext=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.wraptext=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.wraptext=-1"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.wraptext=-1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.wraptext=-1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(63)  =   "Named:id=33:Normal"
      _StyleDefs(64)  =   ":id=33,.parent=0"
      _StyleDefs(65)  =   "Named:id=34:Heading"
      _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   ":id=34,.wraptext=-1"
      _StyleDefs(68)  =   "Named:id=35:Footing"
      _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   "Named:id=36:Selected"
      _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=37:Caption"
      _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(74)  =   "Named:id=38:HighlightRow"
      _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=39:EvenRow"
      _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(78)  =   "Named:id=40:OddRow"
      _StyleDefs(79)  =   ":id=40,.parent=33"
      _StyleDefs(80)  =   "Named:id=41:RecordSelector"
      _StyleDefs(81)  =   ":id=41,.parent=34"
      _StyleDefs(82)  =   "Named:id=42:FilterBar"
      _StyleDefs(83)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmBusqueda
'    Project    : Contabilidad
'
'    Description: formulario de busqueda de datos (nueva version)
'--------------------------------------------------------------------------------
Option Explicit
Dim nUltColVis As Variant
Dim nFirstRow As Variant

Dim Col As TrueOleDBGrid70.Column
Dim COLS As TrueOleDBGrid70.Columns

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Activate
' Description:       Evento que se ejecuta al activar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Activate()
    tdbgListado.FilterActive = True
    Call tdbgListado_FilterChange
    '*** DETERMINANDO LA ULTIMA COLUMNA VISIBLE ***
    Dim i As Integer
    If IsEmpty(nUltColVis) Then
        For i = 0 To tdbgListado.Columns.Count - 1
            If tdbgListado.Columns(i).Visible Then nUltColVis = i
        Next
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    gsCodigo = ""
    gsDetalle = ""
    gsCampo3 = ""
    gsCampo4 = ""
    gsCampo5 = ""
    gsCampo6 = ""
    
    Call Centrar_form(Me)
    nFirstRow = 1
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    tdbgListado.DataSource = Nothing
    Set frmBusqueda = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgListado_DblClick
' Description:       Evento que se ejecuta al hacer doble clic en la lista de busquedas
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbgListado_DblClick()
    Call tdbgListado_KeyDown(vbKeyReturn, 0)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgListado_FilterChange
' Description:       Evento que se ejecuta al cambiar el filtro de la lista
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbgListado_FilterChange()
    On Error GoTo ErrHandler
    Set COLS = tdbgListado.Columns
    Dim c As Integer
    c = tdbgListado.Col
    If UCase(COLS(c).FilterText) = "" Then Exit Sub
    '*** CONVERTIR A MAYUSCULAS
    COLS(c).FilterText = UCase(COLS(c).FilterText)
    grstBusqueda.Filter = GetFilter()
    '*** SI EL FILTRO NO DEVUELVE FILAS RETORNAR AL FILTRO ANTERIOR
    If GetRsRecordCount(grstBusqueda) = 0 Then
        COLS(c).FilterText = tdbgListado.Columns(c).Tag
        grstBusqueda.Filter = GetFilter()
    End If
    tdbgListado.Columns(c).Tag = COLS(c).FilterText
    tdbgListado.Col = c
    tdbgListado.EditActive = True
    Exit Sub
ErrHandler:
    Call ClearFilter
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       GetFilter
' Description:       Procedimiento que obtieneel filtro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetFilter() As String
On Error GoTo ErrHandler
    Dim tmp As String, sBuscar As String
    Dim n As Integer
    For Each Col In COLS
        sBuscar = Replace(Trim(Col.FilterText), "'", "''")      ' CAMBIANDO LAS COMILLAS SIMPLES
        If sBuscar <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            If Col.NumberFormat = "Standard" Then
                tmp = tmp & Col.DataField & " = " & sBuscar
            Else
                tmp = tmp & Col.DataField & " LIKE '%" & sBuscar & "%'"
            End If
        End If
    Next Col
    GetFilter = tmp
    Exit Function
ErrHandler:
    GetFilter = ""
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ClearFilter
' Description:       Procedimiento que limpia el filtro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub ClearFilter()
On Error GoTo ErrHandler
    For Each Col In tdbgListado.Columns
        Col.FilterText = ""
        Col.Tag = ""
    Next Col
    grstBusqueda.Filter = adFilterNone
    Exit Sub
ErrHandler:
    grstBusqueda.Filter = adFilterNone
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgListado_HeadClick
' Description:       Evento que se ejecuta al hacer clic en la cabecera de la grilla
'
' Parameters :       ColIndex (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgListado_HeadClick(ByVal ColIndex As Integer)
    Dim sSort As String
    On Error Resume Next
    sSort = tdbgListado.Columns(ColIndex).DataField
    If sSort = "" Then Exit Sub
    If grstBusqueda.Sort = sSort Then
        grstBusqueda.Sort = sSort + " DESC"
    Else
        grstBusqueda.Sort = sSort
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgListado_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en la grilla
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbgListado_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight And tdbgListado.Col = nUltColVis Then KeyCode = 0: tdbgListado.Col = 0
    If KeyCode = vbKeyLeft And tdbgListado.Col = 0 Then KeyCode = 0: tdbgListado.Col = nUltColVis
    If KeyCode = vbKeyReturn Then
        If tdbgListado.FilterActive Then
            KeyCode = vbKeyDown
        Else
            If GetRsRecordCount(grstBusqueda) > 0 Then
                gsCodigo = tdbgListado.Columns(0).Text
                gsDetalle = tdbgListado.Columns(1).Text
                gsCampo3 = tdbgListado.Columns(2).Text
                gsCampo4 = tdbgListado.Columns(3).Text
                gsCampo5 = tdbgListado.Columns(4).Text
                gsCampo6 = tdbgListado.Columns(5).Text
                
                Unload Me
            End If
        End If
    End If
    If KeyCode = vbKeyEscape Then
        gsCodigo = ""
        gsDetalle = ""
        gsCampo3 = ""
        gsCampo4 = ""
        gsCampo5 = ""
        gsCampo6 = ""
        Unload Me
    End If
End Sub

