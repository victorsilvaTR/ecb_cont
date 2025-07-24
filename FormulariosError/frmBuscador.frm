VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBuscador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   4125
   ClientTop       =   3600
   ClientWidth     =   6105
   Icon            =   "frmBuscador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optBusqueda 
      Caption         =   "Por Descripción"
      Height          =   195
      Index           =   1
      Left            =   2955
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optBusqueda 
      Caption         =   "Por Código"
      Height          =   195
      Index           =   0
      Left            =   1515
      TabIndex        =   4
      Top             =   5040
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TDBText6Ctl.TDBText txtDescripcion 
      Height          =   300
      Left            =   1350
      TabIndex        =   1
      Top             =   45
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   529
      Caption         =   "frmBuscador.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBuscador.frx":0F36
      Key             =   "frmBuscador.frx":0F54
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
      MaxLength       =   50
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
   Begin TDBText6Ctl.TDBText txtCodigo 
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   45
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   529
      Caption         =   "frmBuscador.frx":0F98
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBuscador.frx":1004
      Key             =   "frmBuscador.frx":1022
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
      Format          =   "a@"
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   13
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
   Begin TrueOleDBGrid70.TDBGrid TDBGTabla 
      Height          =   4485
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7911
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CODIGO"
      Columns(0).DataField=   "codigo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DESCRIPCION"
      Columns(1).DataField=   "descripcion"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nivel"
      Columns(2).DataField=   "pla_cTitulo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cod Tipo"
      Columns(3).DataField=   "Mon_cCodigo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=79"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=79"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(23)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=16,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Arial"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText txtRuc 
      Height          =   300
      Left            =   4635
      TabIndex        =   2
      Top             =   45
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   529
      Caption         =   "frmBuscador.frx":1066
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBuscador.frx":10D2
      Key             =   "frmBuscador.frx":10F0
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
      MaxLength       =   50
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
End
Attribute VB_Name = "frmBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmBuscador
'    Project    : Contabilidad
'
'    Description: Formulario de Busquedas
'--------------------------------------------------------------------------------
Option Explicit
Public frmOrigen As Form
Public tabla As String
Public enUso As Boolean
Public auxiliar As String
Public NombreOrigen As String
Public NombreBuscador As String
Public nDigitos As Integer

Dim rsDatos As ADODB.Recordset
Dim cambio As Boolean
Dim gsDigitosCtaDetalle   As Integer

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Cerrar
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub Cerrar()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Activate
' Description:       Evento que se ejecuta al activarse el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Activate()
    If tabla = "CuentasN" Or tabla = "CuentasB" Or tabla = "Cuentas" Or tabla = "CuentasFilt" Or tabla = "EntidadesR" Then
       On Error GoTo serror
       Me.txtCodigo.Text = auxiliar
       txtCodigo.SelStart = Len(txtCodigo.Text)
       pSetFocus Me.txtCodigo
    End If
    
    Exit Sub
serror:
       Me.txtCodigo.Text = ""
       txtCodigo.SelStart = Len(txtCodigo.Text)
       pSetFocus Me.txtCodigo

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
   Dim i As Integer
   gsDigitosCtaDetalle = NE(BuscaValorEnOp("054"))
   
   txtRuc.Visible = False
   gsKeyPressF1 = False
    
   Call Centrar_form(Me)
   Dim sql As String

   enUso = True
   Select Case tabla

   
       Case "TipoCambio"
            Me.Caption = "Busqueda de Tipo de Cambio"
            TDBGTabla.Columns(0).Caption = "Tipo de Cambio"
            
            sql = "SELECT Tab_cCodigo AS Codigo, Tab_cDescripCampo AS Descripcion " & _
                  "From tabla " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Tab_cTabla = '026'"
       
       Case "EntidadesR"
       
            Me.Caption = "Busqueda de Entidades"
            TDBGTabla.Columns(0).Caption = "Codigo"
            
            sql = "SELECT Ent_cCodEntidad,Ent_cPersona,Ent_nRuc,Ten_cTipoEntidad AS Descripcion " & _
                  "From CNM_ENTIDAD " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Ten_cTipoEntidad = 'R'"
                  
        Case "Cuentas0-31"

                    Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "Cuenta"
                        
            sql = "select pla_cCuentaContable, Pla_cNombreCuenta from CNM_PLAN_CTA " & _
                  "where left(pla_cCuentaContable,2)='31' and Pan_cAnio='" & gsAnio & "' and Emp_cCodigo='" & gsEmpresa & "'"
       
       Case "Cuentas10"
            Me.Caption = "Busqueda de Cuentas "
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From dbo.CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND " & _
                  "Pla_cEstado = 'A' AND PLA_CTITULO<>'S' and Left(Pla_cCuentaContable,2)='10' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "Cuentas39"
            Me.Caption = "Busqueda de Cuentas "
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From dbo.CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND " & _
                  "Pla_cEstado = 'A' AND PLA_CTITULO<>'S' and Left(Pla_cCuentaContable,2)='39' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "Cuentas", "CuentasN", "CuentasB" 'solo cuentas detalle
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From dbo.CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND " & _
                  "Pla_cEstado = 'A' AND PLA_CTITULO<>'S' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "CuentasND" 'considerando numero de digitos
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' AND " & _
                  "len(Pla_cCuentaContable) = " & nDigitos & " and Pla_cCuentaContable Like '" & auxiliar & "%' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "CuentasDestino"
            
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' AND " & _
                  "len(Pla_cCuentaContable) = " & gsDigitosCtaDetalle & " and left(Pla_cCuentaContable,1)<> left('" & auxiliar & "',1) " & _
                  "ORDER BY Pla_cCuentaContable "
        
            Me.Caption = "Busqueda de Cuentas"
                  
       Case "CuentasNo2D" 'cuentas mayores a dos digitos para filtro de flujo de efectivo
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' AND " & _
                  "len(Pla_cCuentaContable) > 2 and Pla_cCuentaContable Like '" & auxiliar & "%' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "CuentasFilt"  'cuentas detalle y de titulo
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From dbo.CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' and " & _
                  "Pla_cCuentaContable Like '" & auxiliar & "%' " & _
                  "ORDER BY Pla_cCuentaContable "
       
'       Case "CuentasB"  'cuentas detalle
'            Me.Caption = "Busqueda de Cuentas"
'            TDBGTabla.Columns(0).Caption = "CUENTA"
'            Sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
'                  "From dbo.CNM_PLAN_CTA " & _
'                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' and " & _
'                  "Pla_cCuentaContable Like '" & auxiliar & "%' And Pla_cTitulo='N'" & _
'                  "ORDER BY Pla_cCuentaContable "
      
       Case "CuentasFiltCaja" 'solo cuentas de caja
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' and " & _
                  "Pla_cCuentaContable Like '10%' and PLA_CTITULO<>'S' and Pla_cCuentaContable Like '" & auxiliar & "%' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "CuentasFiltCajaF0101" 'solo cuentas de caja FORMATO 0101
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' and " & _
                  "Pla_cCuentaContable Like '10%' and PLA_CTITULO<>'S' AND LEFT(Pla_cCuentaContable,3)<='102' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "CuentasFiltCajaF0102" 'solo cuentas de caja FORMATO 0102
            Me.Caption = "Busqueda de Cuentas"
            TDBGTabla.Columns(0).Caption = "CUENTA"
            sql = "SELECT Pla_cCuentaContable as codigo, Pla_cNombreCuenta as descripcion, pla_cTitulo " & _
                  "From CNM_PLAN_CTA " & _
                  "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cEstado = 'A' and " & _
                  "Pla_cCuentaContable Like '10%' and PLA_CTITULO<>'S' AND LEFT(Pla_cCuentaContable,3)>='103' " & _
                  "ORDER BY Pla_cCuentaContable "
       
       Case "TipoDocumento"
            Me.Caption = "Busqueda de Tipo de Documento"
            sql = "spCn_ConsultaTipDocsLibro 'SEL_DOCS_ALL','" & gsEmpresa & "','" & gsAnio & "'"
            
        Case "TipoDocumentoAsiento"
            Me.Caption = "Busqueda de Tipo de Documento"
            sql = "spCn_ConsultaTipDocsLibro 'SEL_DOCS_ALL_LIBRO','" & gsEmpresa & _
                    "', '" & gsAnio & "', '" & auxiliar & "',''"
       
       Case "Libros"
            Me.Caption = "Busqueda de Tipo de Libros de Operaciones"
            sql = "SELECT Lib_cTipoLibro as codigo, Lib_cDescripcion as descripcion " & _
                  "FROM CNT_LIBRO_OPERA  " & _
                  "WHERE  Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
                  "ORDER BY Lib_cDescripcion "
       
       Case "CentroCostoN", "CentroCosto"
            Me.Caption = "Busqueda de Centro de Costo"
'            sql = "SELECT Cos_cCodigo As codigo, Cos_cDescripcion As descripcion, Cos_cTitulo As pla_cTitulo " & _
'                  "FROM CNT_CENTRO_COSTO " & _
'                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' And cos_cEstado = 'A' and " & _
'                  "Pan_cAnio='" & gsAnio & "' and cos_ctitulo='N' AND cos_cStitulo='N' " & _
'                  "Order By Cos_cCodigo "
                  
            sql = "EXEC spCNT_CENTRO_COSTO 'BUSAR_NIVEL_F1_CONTA', '" & gsEmpresa & "','" & gsAnio & "'"
       
       Case "CentroCostoPres"
            Me.Caption = "Busqueda de Centro de Costo"
'            Sql = "SELECT Cos_cCodigo As codigo, Cos_cDescripcion As descripcion  " & _
'                  "FROM CNT_CENTRO_COSTO " & _
'                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' And cos_cEstado = 'A' and Pan_cAnio='" & gsAnio & "' " & _
'                  "Order By Cos_cCodigo "
            sql = "EXEC spCNT_CENTRO_COSTO 'BUSAR_NIVEL_F1_CONTA', '" & gsEmpresa & "','" & gsAnio & "'"
            
       Case "Balance"
            Me.Caption = "Busqueda de Balance"
            sql = "Select Ppa_cNumPlantilla as codigo, Ppa_cNombre as descripcion, Ppa_cTitulo as pla_cTitulo " & _
                  "From CNA_TIPO_PLANTILLA  " & _
                  "Where Emp_cCodigo='" & gsEmpresa & "' And Ppa_cTitulo='N' And Ppa_cNombre<>'' And Ppa_cTipoPlantilla='BGE' " & _
                  " AND Pan_cAnio = '" & gsAnio & "' " & _
                  "Order by Ppa_cNumPlantilla"
       
       Case "Funcion"
            Me.Caption = "Busqueda de Resultado por Función"
            sql = "Select Ppa_cNumPlantilla as codigo, Ppa_cNombre as descripcion, Ppa_cTitulo as pla_cTitulo " & _
                  "From CNA_TIPO_PLANTILLA  " & _
                  "Where Emp_cCodigo='" & gsEmpresa & "' And Ppa_cTitulo='N' And Ppa_cNombre<>'' And Ppa_cTipoPlantilla='FUN' " & _
                  " AND Pan_cAnio = '" & gsAnio & "' " & _
                  "Order by Ppa_cNumPlantilla"
       
       Case "Naturaleza"
            Me.Caption = "Busqueda de Resultado por Naturaleza"
            sql = "Select Ppa_cNumPlantilla as codigo, Ppa_cNombre as descripcion, Ppa_cTitulo as pla_cTitulo " & _
                  "From CNA_TIPO_PLANTILLA  " & _
                  "Where Emp_cCodigo='" & gsEmpresa & "' And Ppa_cTitulo='N' And Ppa_cNombre<>'' And Ppa_cTipoPlantilla='NAT' " & _
                  " AND Pan_cAnio = '" & gsAnio & "' " & _
                  "Order by Ppa_cNumPlantilla"
       
       Case "CuentaRatio"
            Me.Caption = "Busqueda de Cuentas de Ratio"
            sql = "SELECT Ind_cCodCuenta as codigo, Ind_cDescripcion as descripcion " & _
                  "FROM CNT_CUENTA_INDI  " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
                  "ORDER BY Ind_cDescripcion "
       
       Case "Entidad", "EntidadB"
            Me.Caption = "Busqueda de Entidades"
            TDBGTabla.Columns(0).Width = 700
            TDBGTabla.Columns(1).Width = 3100
            
            TDBGTabla.Columns(2).Visible = True
            TDBGTabla.Columns(2).Caption = "RUC"
            TDBGTabla.Columns(2).Width = 1500
            
            TDBGTabla.Columns(3).Visible = True
            TDBGTabla.Columns(3).Caption = "T"
            TDBGTabla.Columns(3).Width = 100

            txtDescripcion.Width = 3300
            txtRuc.Visible = True
            If Trim(auxiliar) = "" Then
                sql = "SELECT ent.Ent_cCodEntidad ,CASE TE.Ten_cPlame WHEN '1' THEN ENT.Ent_cApaterno+' '+ENT.Ent_cAmaterno+' '+ENT.Ent_cNombres " & _
                      "ELSE  ENT.Ent_cPersona END as 'Ent_cPersona', ent.Ent_nRuc , ent.Ten_CTipoEntidad as pla_cTitulo " & _
                      "FROM CNM_ENTIDAD ent inner join CNT_ENTIDAD te WITH(NOLOCK) ON ENT.Emp_cCodigo = TE.Emp_cCodigo AND ENT.Ten_cTipoEntidad = TE.Ten_cTipoEntidad  " & _
                      "WHERE ent.Emp_cCodigo = '" & gsEmpresa & "' AND ent.Ent_cEstado = 'A' " & _
                      "ORDER BY Ent_cPersona "
            Else
                sql = "SELECT ent.Ent_cCodEntidad ,CASE TE.Ten_cPlame WHEN '1' THEN ENT.Ent_cApaterno+' '+ENT.Ent_cAmaterno+' '+ENT.Ent_cNombres " & _
                      "ELSE  ENT.Ent_cPersona END as 'Ent_cPersona', ent.Ent_nRuc , ent.Ten_CTipoEntidad as pla_cTitulo  " & _
                      "FROM CNM_ENTIDAD ent inner join CNT_ENTIDAD te WITH(NOLOCK) ON ENT.Emp_cCodigo = TE.Emp_cCodigo AND ENT.Ten_cTipoEntidad = TE.Ten_cTipoEntidad  " & _
                      "WHERE ent.Emp_cCodigo = '" & gsEmpresa & "' AND ent.Ent_cEstado = 'A' " & _
                      "and ent.Ten_CTipoEntidad =  '" & CE(auxiliar) & "' " & _
                      "ORDER BY ent.Ent_cPersona "
            End If
       Case "Empresas"
            Me.Caption = "Busqueda de Empresas"
            sql = "SELECT Emp_cCodigo as codigo, Emp_cNombreLargo as descripcion,Emp_cNombreCorto,  " & _
                  "Emp_cDireccion, Emp_cNumRuc ,Emp_cTelefono, Emp_cCodSuc " & _
                  "FROM EMPRESA  " & _
                  "WHERE EMP_CCODIGO<>'000'  And (Emp_cEstado <> '*' OR Emp_cEstado IS NULL  )  " & _
                  "ORDER BY Emp_cCodigo "
       
       Case "EmpresasAll"
            Me.Caption = "Busqueda de Empresas"
            sql = "SELECT Emp_cCodigo as codigo, Emp_cNombreLargo as descripcion " & _
                  "FROM EMPRESA  " & _
                  "WHERE Emp_cEstado <> '*' OR (Emp_cEstado IS NULL  AND  EMP_CCODIGO<>'000' ) " & _
                  "ORDER BY Emp_cNombreLargo "
   End Select
   
    '----------------------------------------------
    Call CerrarRecordSet(rsDatos)
    
    Set rsDatos = fRetornaRS(sql)
    '----------------------------------------------
    If Not rsDatos Is Nothing Then
        Set TDBGTabla.DataSource = rsDatos
    End If
    '----------------------------------------------
    On Error GoTo serror
    With TDBGTabla
        .AllowUpdate = False
        .HighlightRowStyle = "HighlightRow"
        
        If Not rsDatos Is Nothing Then
            If rsDatos.RecordCount > 0 Then
                For i = 0 To rsDatos.Fields.Count - 1
                    .Columns(i).DataField = rsDatos.Fields(i).Name
                Next i
                
                'On Error Resume Next
                'txtCodigo.Text = CE(auxiliar)
            End If
        End If
    End With
    '----------------------------------------------
    
    'txtCodigo.SelStart = Len(txtCodigo.Text)
serror:
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       EnviaCodigo
' Description:       Procedimiento que envia el codigo seleccionado al formulario que lo invoca
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub EnviaCodigo()

   If Not rsDatos Is Nothing Then
   If rsDatos.RecordCount > 0 Then
      
      Select Case tabla
      
        Case "CuentasN"
            If Me.TDBGTabla.Columns(2).Value = "S" Then
                Mensajes "Es cuenta de titulo. Verifique seleccion...", vbInformation
                pSetFocus TDBGTabla
                Exit Sub
            End If
        Case "CentroCostoN"
            If Me.TDBGTabla.Columns(2).Value = "S" Then
                Mensajes "Es Centro Costo de titulo. Verifique seleccion...", vbInformation
                pSetFocus TDBGTabla
                Exit Sub
            End If
        Case "Balance"
            If Me.TDBGTabla.Columns(2).Value = "S" Then
                Mensajes "Es titulo. Verifique seleccion...", vbInformation
                pSetFocus TDBGTabla
                Exit Sub
            End If
        Case "Funcion"
            If Me.TDBGTabla.Columns(2).Value = "S" Then
            
                Mensajes "Es titulo. Verifique seleccion...", vbInformation
                pSetFocus TDBGTabla
                Exit Sub
            End If
        Case "Naturaleza"
            If Me.TDBGTabla.Columns(2).Value = "S" Then
                Mensajes "Es titulo. Verifique seleccion...", vbInformation
                Exit Sub
            End If
      End Select
      
    If NombreOrigen = "frmManAsientosContables" And frmMDIConta.BuscaForm("frmManAsientosContables") = True Then
    
        frmManAsientosContables.Enabled = True
        frmManAsientosContables.RecibirDatos tabla, Me.TDBGTabla.Columns(0).Value, Me.TDBGTabla.Columns(1).Value, Me.TDBGTabla.Columns(3).Value
        gsKeyPressF1 = True
        frmManAsientosContables.LimpiaCeldasInactivas
    Else
        If NombreOrigen = "frmManEmpresas" Then
            frmOrigen.Enabled = True
            frmOrigen.RecibirDatos tabla, CE(rsDatos.Fields(0).Value), _
                                          CE(rsDatos.Fields(1).Value), _
                                          CE(rsDatos.Fields(2).Value), _
                                          CE(rsDatos.Fields(3).Value), _
                                          CE(rsDatos.Fields(4).Value), _
                                          CE(rsDatos.Fields(5).Value), _
                                          CE(rsDatos.Fields(6).Value)
'        ElseIf NombreOrigen = "FrmManRegAcc" Then
'            frmOrigen.Enabled = True
'            frmOrigen.RecibirDatos tabla, CE(rsDatos.Fields(0).Value)
        Else
            frmOrigen.Enabled = True
'            Set frmOrigen = FrmManRegAcc


If Me.TDBGTabla.Columns(3).Value <> "" And NombreOrigen = "FrmManRegAcc" Then
    frmOrigen.RecibirDatos tabla, Me.TDBGTabla.Columns(0).Value, Me.TDBGTabla.Columns(1).Value, Me.TDBGTabla.Columns(2).Value, Me.TDBGTabla.Columns(3).Value
Else
    frmOrigen.RecibirDatos tabla, Me.TDBGTabla.Columns(0).Value, Me.TDBGTabla.Columns(1).Value, Me.TDBGTabla.Columns(2).Value
End If


            
        End If
    End If
        
      gsKeyCodePress = False
      Unload Me
   Else
      Mensajes "Código no existe, digite correctamente... ", vbOKOnly + vbInformation
      If Me.optBusqueda(0).Value = True Then pSetFocus txtCodigo
      If Me.optBusqueda(1).Value = True Then pSetFocus txtDescripcion
   End If
   
   End If
   
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       FiltrarADC
' Description:       Procedimiento que filtra el recordset del formulario de busqueda
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub FiltrarADC()
   Dim cadena As String
   Dim filtros(2) As String
   Dim i As Integer
   
   On Error GoTo serror
   cadena = ""
   If CE(Me.txtCodigo.Text) <> "" Then filtros(0) = TDBGTabla.Columns(0).DataField & " like '" & txtCodigo & "*'"
   If CE(Me.txtDescripcion.Text) <> "" Then filtros(1) = TDBGTabla.Columns(1).DataField & " like '*" & txtDescripcion & "*'"
   If CE(Me.txtRuc.Text) <> "" Then filtros(2) = TDBGTabla.Columns(2).DataField & " like '*" & txtRuc & "*'"
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
   If CE(cadena) <> "" Then
      rsDatos.Filter = cadena
   Else
      rsDatos.Filter = 0
   End If
   
   
   Exit Sub
serror:
    If Not rsDatos Is Nothing Then
        rsDatos.Filter = 0
    End If
'   If cambio = True Then
'      rsDatos.Sort = "descripcion, codigo"
'   Else
'      rsDatos.Sort = "codigo, descripcion"
'   End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

    If NombreOrigen = "frmManAsientosContables" And frmMDIConta.BuscaForm("frmManAsientosContables") = True Then
        frmManAsientosContables.Enabled = True
    Else
        If Not frmOrigen Is Nothing Then
            frmOrigen.Enabled = True
        End If
    End If
    
    enUso = False
    gsKeyCodePress = False
    
    Set frmOrigen = Nothing
    Set TDBGTabla.DataSource = Nothing
    
    Call CerrarRecordSet(rsDatos)
    
    Set frmBuscador = Nothing
        
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       TDBGTabla_DblClick
' Description:       Evento que se ejecuta al hacer doble click en la grilla del buscador
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub TDBGTabla_DblClick()
   Call EnviaCodigo
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       TDBGTabla_KeyUp
' Description:       Evento que se ejecuta al presionar una tecla
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub TDBGTabla_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call TDBGTabla_DblClick   ' EnviaCodigo
    

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtCodigo_Change
' Description:       Evento que se ejecuta al cambiar el codigo de busqueda
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtCodigo_Change()
   cambio = False
   FiltrarADC
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtCodigo_KeyDown
' Description:       Evento que se ejecuta alpresionar una tecla en el codigo de busqueda
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo serror
    Dim i As Integer
   If KeyCode = 40 Then
      TDBGTabla.MoveNext
      If TDBGTabla.EOF Then TDBGTabla.MoveLast
      KeyCode = 0
   End If
   If KeyCode = 38 Then
      TDBGTabla.MovePrevious
      If TDBGTabla.BOF Then TDBGTabla.MoveFirst
      KeyCode = 0
   End If
    
    If KeyCode = vbKeyRight Then
        pSetFocus txtDescripcion
        KeyCode = 0
    End If

    If KeyCode = vbKeyLeft Then
        If txtRuc.Visible = False Then pSetFocus txtDescripcion
        If txtRuc.Visible = True Then pSetFocus txtRuc
        KeyCode = 0
    End If
   
    If KeyCode = 34 Then
        For i = 0 To 11
            TDBGTabla.MoveNext
        Next
        If TDBGTabla.EOF Then TDBGTabla.MoveLast
        KeyCode = 0
        'TDBGTabla.MoveRelative (6)
    End If
    If KeyCode = 33 Then
        For i = 0 To 11
            TDBGTabla.MovePrevious
        Next
        If TDBGTabla.BOF Then TDBGTabla.MoveFirst
        KeyCode = 0
    End If
    Exit Sub
serror:
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtCodigo_KeyPress
' Description:       Evento que se ejecuta al presionaruna tecla en el codigo de busqueda
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      EnviaCodigo
   End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtDescripcion_Change
' Description:       Evento que se ejecuta al cambiar la descripcion de busqueda
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtDescripcion_Change()
   cambio = True
   
    If gsKey = 219 Then
        txtDescripcion = Replace(txtDescripcion, "'", "")
        txtDescripcion.SelStart = Len(txtDescripcion)
    End If
    
   FiltrarADC
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtDescripcion_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en la descripcion de busqueda
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    Dim i As Integer
    
    If KeyCode = 40 Then
       TDBGTabla.MoveNext
       If TDBGTabla.EOF Then TDBGTabla.MoveLast
       KeyCode = 0
    End If
    If KeyCode = 38 Then
       TDBGTabla.MovePrevious
       If TDBGTabla.BOF Then TDBGTabla.MoveFirst
       KeyCode = 0
    End If
    
    If KeyCode = vbKeyRight Then
        If txtRuc.Visible = False Then pSetFocus txtCodigo
        If txtRuc.Visible = True Then pSetFocus txtRuc
        KeyCode = 0
    End If

    If KeyCode = vbKeyLeft Then
        pSetFocus txtCodigo
        KeyCode = 0
    End If
    
    
    If KeyCode = 34 Then
        For i = 0 To 11
            TDBGTabla.MoveNext
        Next
        If TDBGTabla.EOF Then TDBGTabla.MoveLast
        KeyCode = 0
        'TDBGTabla.MoveRelative (6)
    End If
    If KeyCode = 33 Then
        For i = 0 To 11
            TDBGTabla.MovePrevious
        Next
        If TDBGTabla.BOF Then TDBGTabla.MoveFirst
        KeyCode = 0
    End If
    
   
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtDescripcion_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el campo de descripcion
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then EnviaCodigo
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtRuc_Change
' Description:       Evento que se ejecuta al cambiar el filtro del ruc
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub txtRuc_Change()
   FiltrarADC
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtRuc_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el campo del ruc
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub txtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    Dim i As Integer

    If KeyCode = 40 Then
       TDBGTabla.MoveNext
       If TDBGTabla.EOF Then TDBGTabla.MoveLast
       KeyCode = 0
    End If
    If KeyCode = 38 Then
       TDBGTabla.MovePrevious
       If TDBGTabla.BOF Then TDBGTabla.MoveFirst
       KeyCode = 0
    End If
    
    
    If KeyCode = vbKeyRight Then
        pSetFocus txtCodigo
        KeyCode = 0
    End If

    If KeyCode = vbKeyLeft Then
        pSetFocus txtDescripcion
        KeyCode = 0
    End If
    
    
    If KeyCode = 34 Then
        For i = 0 To 11
            TDBGTabla.MoveNext
        Next
        If TDBGTabla.EOF Then TDBGTabla.MoveLast
        KeyCode = 0
        'TDBGTabla.MoveRelative (6)
    End If
    If KeyCode = 33 Then
        For i = 0 To 11
            TDBGTabla.MovePrevious
        Next
        If TDBGTabla.BOF Then TDBGTabla.MoveFirst
        KeyCode = 0
    End If


End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       txtRuc_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el campo del ruc
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub txtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then EnviaCodigo
End Sub
