VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcReindex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reindexación de la Base de datos"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmPrcReindex.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   5070
   Begin VB.Frame fraTodo 
      Height          =   5220
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4935
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4605
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   135
         TabIndex        =   1
         Top             =   4365
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   2565
         TabIndex        =   4
         Top             =   4680
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdproc 
         Height          =   435
         Left            =   765
         TabIndex        =   3
         Top             =   4680
         Width           =   1665
         Caption         =   " Reindexar"
         PicturePosition =   327683
         Size            =   "2937;767"
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
      TabIndex        =   5
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcReindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdproc_Click()
Dim X As Integer
    On Error GoTo serror
    If MsgBox("Ningún usuario debe de estar utilizando el Sistema, Continua", vbYesNo + vbInformation, gsNombreModulo) = vbYes Then
        cmdproc.Enabled = False
        cmdSalir.Enabled = False
        Set cn = New ADODB.Connection
        cn.ConnectionString = gsCadenaConexion
        cn.Open
        ProgressBar1.Min = 0
        ProgressBar1.Max = List1.ListCount
        Screen.MousePointer = vbHourglass
        
        Call EscribirLog("Iniciando la indexacion y defragmentación de la base de datos", Me.Name)
        
        For X = 0 To List1.ListCount - 1
            List1.ListIndex = X
            DoEvents
            On Error Resume Next
            cn.Execute "DBCC DBREINDEX (" & List1.List(X) & ",'', 0)"
            ProgressBar1.Value = X
        Next
        
        List1.ListIndex = 0
        For X = 0 To List1.ListCount - 1
            Me.Caption = "Defragmentación de la Base de datos"
            List1.ListIndex = X
            DoEvents
            On Error Resume Next
            cn.Execute "DBCC INDEXDEFRAG (" & gsBD & "," & List1.List(X) & ")"
            ProgressBar1.Value = X
        Next
        
        Me.Caption = "Reindexación de la Base de datos"
        MsgBox "El proceso se realizo satisfactoriamente", vbInformation, gsNombreModulo
        
        Call EscribirLog("Finalizo la indexacion de la base de datos", Me.Name)
        
    End If
    ProgressBar1.Value = ProgressBar1.Max
    Screen.MousePointer = vbNormal
    cmdproc.Enabled = True
    cmdSalir.Enabled = True
    Exit Sub
serror:
    Call EscribirLog("Error a indexar la base de datos", Me.Name)
    Mensajes Err.Description
End Sub

Sub CargaTablas()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    sqlSp = "select name from sysobjects where xtype = 'U' and ( left(name,2)='CN' or left(name,2)='SG' ) order by name "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 1 Then
        Do While Not rsArreglo.EOF
           List1.AddItem CE(rsArreglo!Name)
           rsArreglo.MoveNext
        Loop
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    


'List1.AddItem "CNA_CONCEPTO_CUENTA"
'List1.AddItem "CNA_CONCEPTO_IMP"
'List1.AddItem "CNA_CTAS_CONDESTINO"
'List1.AddItem "CNA_IMPRESION_CUENTA"
'List1.AddItem "CNA_TIPO_PLANTILLA"
'List1.AddItem "CNC_ASIENTO_VOUCHER"
'List1.AddItem "CND_ASIENTO_COA"
'List1.AddItem "CNA_CONCEPTO_CUENTA"
'List1.AddItem "CND_ASIENTO_VOUCHER"
'List1.AddItem "CND_BALANCE_SUNAT"
'List1.AddItem "CND_CONFIG_OPERA"
'
'List1.AddItem "CND_CONFIG_OPERA"
'
'List1.AddItem "CND_CUENTA_DIST"
'List1.AddItem "CND_SALDOS"
'List1.AddItem "CNM_CUENTA_BANCO"
'List1.AddItem "CNM_ENTIDAD"
'List1.AddItem "CNM_MOV_CHEQUE"
'
'List1.AddItem "CNM_PLAN_CTA"
'List1.AddItem "CNM_SALDOS_RATIOS"
'List1.AddItem "CNT_ANIO"
'List1.AddItem "CNT_ASIENTO_LIBRO"
'List1.AddItem "CNT_BANCO"
'List1.AddItem "CNT_CENTRO_COSTO"
'List1.AddItem "CNT_CIERRE"
'
'List1.AddItem "CNT_COMPRA_IGV"
'List1.AddItem "CNT_CONFIG_LIBROS"
'List1.AddItem "CNT_CONFIG_OPERA"
'List1.AddItem "CNT_CUENTA_INDI"
'List1.AddItem "CNT_ENTIDAD"
'List1.AddItem "CNT_ENTIDAD_DOCU"
'List1.AddItem "CNT_INDICADORES"
'List1.AddItem "CNT_LIBRO_OPERA"
'List1.AddItem "CNT_LIBRO_TIPODOC"
'List1.AddItem "CNT_OPERA_ESTADO"
'List1.AddItem "CNT_PERIODO"
'List1.AddItem "CNT_PLAN_BALANCE"
'List1.AddItem "CNT_PLAN_GESTION"
'List1.AddItem "CNT_TIPO_CAMBIO"
'List1.AddItem "CNT_TIPO_MONEDA"
'List1.AddItem "CNT_TIPODOC"
'List1.AddItem "PRM_MARCO_PRES"
'List1.AddItem "SGM_ACCESOS"
'List1.AddItem "SGM_CABPLAN"
'List1.AddItem "SGM_DETPLAN"
'List1.AddItem "SGM_OPMENU"
'List1.AddItem "SGM_PERFIL"
'List1.AddItem "SGM_SOFTWARE"
'List1.AddItem "SGM_USUARIOS"
'List1.AddItem "SIS_RUTAS"
End Sub


Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    Call CargaTablas
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdproc.Enabled = False
        
    Else
        Me.cmdproc.Enabled = True
        
    End If
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
