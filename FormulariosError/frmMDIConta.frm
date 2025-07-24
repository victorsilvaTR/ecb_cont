VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIConta 
   BackColor       =   &H80000004&
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   -3675
   ClientWidth     =   13800
   Icon            =   "frmMDIConta.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFondo1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      Picture         =   "frmMDIConta.frx":0ECA
      ScaleHeight     =   495
      ScaleWidth      =   13740
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   13800
   End
   Begin VB.PictureBox picFondo2 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      Picture         =   "frmMDIConta.frx":1455B
      ScaleHeight     =   555
      ScaleWidth      =   13740
      TabIndex        =   4
      Top             =   915
      Visible         =   0   'False
      Width           =   13800
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   13800
      TabIndex        =   2
      Top             =   5160
      Width           =   13800
      Begin MSComctlLib.TabStrip tsForms 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   582
         MultiRow        =   -1  'True
         Style           =   2
         ImageList       =   "ImgBarraVentanas"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ventanas 01"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte 01"
               ImageVarType    =   2
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8490
      Top             =   1695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":27BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":28186
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":28720
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":28CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":29254
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":297EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":29D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2A322
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2A8BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2AE56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2B3F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2B98A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2BF24
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2C4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2CBD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMdi 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5445
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   176
            Picture         =   "frmMDIConta.frx":2D1FC
            Object.ToolTipText     =   " Usuario activo del sistema "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17163
            MinWidth        =   7056
            Picture         =   "frmMDIConta.frx":2D796
            Object.ToolTipText     =   " Periodo y Año actual del Sistema "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   353
            Picture         =   "frmMDIConta.frx":2DD30
            Object.ToolTipText     =   " Codigo de Empresa "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2E2CA
            Object.ToolTipText     =   " Servidor conectado "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "frmMDIConta.frx":2E864
            Object.ToolTipText     =   " Base de Datos "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13800
      _ExtentX        =   24342
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Plan de cuentas"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tipo de Cambio"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Entidades"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Registro de Asientos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configuracion de Operaciones y Libros vs Cuenta"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Organizar ventanas como Cascada"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Organizar ventanas como Mosaico"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Organizar ventanas horizontalmente"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio de Año de Trabajo / Empresa "
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar de Usuario y Empresa"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   14
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImgBarraVentanas 
      Left            =   6495
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2EDFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIConta.frx":2F398
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTablas 
      Caption         =   "&Configuración"
      Begin VB.Menu mnuPlanCuentas 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu mnuGuion01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCentroCostos 
         Caption         =   "Centro de Costos"
      End
      Begin VB.Menu mnuTipoLibro 
         Caption         =   "Libros"
      End
      Begin VB.Menu mnuEntidades 
         Caption         =   "Entidades"
         Begin VB.Menu mnuTipoEntidad 
            Caption         =   "Tipos de Entidad"
         End
         Begin VB.Menu mnuEntDoc 
            Caption         =   "Entidad - Documento"
         End
         Begin VB.Menu mnuEntidadesMnt 
            Caption         =   "Entidades"
         End
      End
      Begin VB.Menu mnuTipoDocumento 
         Caption         =   "Tipos de Documento"
      End
      Begin VB.Menu mnuTablasMnt 
         Caption         =   "Tablas Secundarias"
      End
      Begin VB.Menu mnuDepA 
         Caption         =   "Dependencias Aduanera"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGuion02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoMoneda 
         Caption         =   "Tipo de Moneda"
      End
      Begin VB.Menu mnuTipoCambio 
         Caption         =   "Tipo de Cambio"
      End
      Begin VB.Menu mnuGuion03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu mnuCuentaCorriente 
         Caption         =   "Cuenta Corriente"
      End
      Begin VB.Menu mnuTiposAsiento 
         Caption         =   "Asientos Tipo"
      End
      Begin VB.Menu mnuPlantillaBalanceGeneral 
         Caption         =   "Plantillas de EEFF"
      End
      Begin VB.Menu mnuPatrimonio 
         Caption         =   "Estado de Cambios en el Patrimonio"
      End
      Begin VB.Menu mnuParamIniciales 
         Caption         =   "Parametros Iniciales"
      End
      Begin VB.Menu mnuConfigOperac 
         Caption         =   "Operaciones y Libros - Cuenta"
      End
      Begin VB.Menu mnuConceptosLibros 
         Caption         =   "Conceptos por Libro"
      End
      Begin VB.Menu mnuGuion07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCambioAnio 
         Caption         =   "Cambio de Año de Trabajo / Empresa"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir del Sistema"
      End
   End
   Begin VB.Menu mnuIngresos 
      Caption         =   "&Ingresos"
      Begin VB.Menu MnuRegistroAsientos 
         Caption         =   "Registro de Asientos"
      End
      Begin VB.Menu mnuRegistroPresupuesto 
         Caption         =   "Registro de Presupuesto"
      End
      Begin VB.Menu mnuCapital 
         Caption         =   "Registro de Capital"
      End
      Begin VB.Menu mnuPDB 
         Caption         =   "PDB"
         Begin VB.Menu mnuPDBCompras 
            Caption         =   "Comprobantes de Compras"
         End
         Begin VB.Menu mnuPDBVentas 
            Caption         =   "Comprobantes de Ventas"
         End
         Begin VB.Menu mnuPDBFPago 
            Caption         =   "Formas de Pago"
         End
      End
      Begin VB.Menu mnuGuion44 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPDT0601 
         Caption         =   "PDT 0601 - PLAME"
      End
      Begin VB.Menu mnuGuion45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuValores 
         Caption         =   "Registro de Inversiones"
      End
      Begin VB.Menu mnuMercaderias 
         Caption         =   "Registro de Mercaderias, Existencias y Suministros"
      End
      Begin VB.Menu mnuIntangibles 
         Caption         =   "Registro de Intangibles"
      End
      Begin VB.Menu mnuArrendaFinanciero 
         Caption         =   "Registro de Activos en Arrendamiento Financiero"
      End
   End
   Begin VB.Menu mnu_reg_axiliar 
      Caption         =   "Reg. &Auxiliares"
      Begin VB.Menu mnu_reg_aux_ventas 
         Caption         =   "Registro Auxiliar de Compras y Ventas"
      End
      Begin VB.Menu mnuReporteRegAuxiliares 
         Caption         =   "Reportes de Registros auxiliares"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "&Procesos"
      Begin VB.Menu mnuGenerarDiferenciaCambioMensual 
         Caption         =   "Generar Diferencia de Cambio Mensual"
      End
      Begin VB.Menu mnuGuion30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCierreEjercicio 
         Caption         =   "1) Cierre del Ejercicio (mes de AJUSTE)"
      End
      Begin VB.Menu mnuAsientoCierre 
         Caption         =   "2) Asiento de Cierre Final (mes de CIERRE)"
      End
      Begin VB.Menu mnuGuion31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAsientoApertura 
         Caption         =   "Asiento de Apertura (mes de APERTURA)"
      End
      Begin VB.Menu mnuGuion32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCierreMensual 
         Caption         =   "Bloquear / Desbloquear meses (seguridad)"
      End
   End
   Begin VB.Menu mnuConciliacionBancaria 
      Caption         =   "Conciliación &Bancaria"
      Begin VB.Menu mnuRegistroEstractoBancario 
         Caption         =   "Registro de Extracto Bancario"
      End
      Begin VB.Menu mnuMovimientosBancos 
         Caption         =   "Reporte de Movimientos de Bancos"
      End
      Begin VB.Menu mnuGuion12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeguimientoCheques 
         Caption         =   "Reporte de Cheques Emitidos"
      End
      Begin VB.Menu mnuChequesPendientes 
         Caption         =   "Reporte de Cheques Pendientes"
      End
   End
   Begin VB.Menu mnuConsultasReportes 
      Caption         =   "&Reporte Análisis"
      Begin VB.Menu mnuSaldosNetos 
         Caption         =   "Saldos Netos Mensuales por Cuenta"
      End
      Begin VB.Menu mnuSaldosCuenta 
         Caption         =   "Saldos del Mes"
      End
      Begin VB.Menu mnuGuion17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBalanceComprobacion 
         Caption         =   "Balance de Comprobación de Análisis"
      End
      Begin VB.Menu mnuReporteAsientosEntidades 
         Caption         =   "Reporte de Analísis por Entidades"
      End
      Begin VB.Menu mnuReporteAnalisisLibro 
         Caption         =   "Reporte de Analisis por Libro"
      End
      Begin VB.Menu MnuEGPPF 
         Caption         =   "Estado del Resultado Integral  (Comparativo Mensual)"
      End
      Begin VB.Menu MnuBalGral 
         Caption         =   "Estado de Situación Financiera (Comparativo Mensual)"
      End
      Begin VB.Menu mnuGuion19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesCentroCosto 
         Caption         =   "Reportes por Centro de Costos"
         Begin VB.Menu mnuAsientosCentroCosto 
            Caption         =   "Analisis por Centro de Costo"
         End
         Begin VB.Menu mnuResumenCentroCosto 
            Caption         =   "Resumen por Centro de Costo"
         End
         Begin VB.Menu mnuResumenCentroCostoMes 
            Caption         =   "Resumen por Centro de Costo Anual"
         End
      End
      Begin VB.Menu mnuReporteEjecucionPresupuesto 
         Caption         =   "Reporte de Ejecución del Presupuesto"
      End
      Begin VB.Menu mnuRepDetracciones 
         Caption         =   "Reporte de Detracciones"
      End
      Begin VB.Menu MnuRepCostCta 
         Caption         =   "Reporte de Costo Por Cuenta"
      End
   End
   Begin VB.Menu mnuRepsunat 
      Caption         =   "Reportes &SUNAT"
      Begin VB.Menu mnuRepInvBal 
         Caption         =   "Saldo de Cuentas de Balance"
         Begin VB.Menu mnuF0302 
            Caption         =   "Detalle Saldo Cuenta 10"
         End
         Begin VB.Menu mnuF0303 
            Caption         =   "Detalle Saldo Cuenta 12"
         End
         Begin VB.Menu mnuF0304 
            Caption         =   "Detalle Saldo Cuenta 14"
         End
         Begin VB.Menu mnuF0305 
            Caption         =   "Detalle Saldo Cuenta 16"
         End
         Begin VB.Menu mnuF0306 
            Caption         =   "Detalle Saldo Cuenta 19"
         End
         Begin VB.Menu mnuF0307_2021 
            Caption         =   "Detalle Saldo Cuenta 2"
         End
         Begin VB.Menu mnuF0308 
            Caption         =   "Detalle Saldo Cuenta 31"
         End
         Begin VB.Menu mnuF0309_33 
            Caption         =   "Detalle Saldo Cuenta 33"
         End
         Begin VB.Menu mnuF0309_34 
            Caption         =   "Detalle Saldo Cuenta 34"
         End
         Begin VB.Menu mnuF0309_35 
            Caption         =   "Detalle Saldo Cuenta 35"
         End
         Begin VB.Menu mnuF0309_38 
            Caption         =   "Detalle Saldo Cuenta 38"
         End
         Begin VB.Menu mnuF0309_39 
            Caption         =   "Detalle Saldo Cuenta 39"
         End
         Begin VB.Menu mnuF0310 
            Caption         =   "Detalle Saldo Cuenta 40"
         End
         Begin VB.Menu mnuF0311 
            Caption         =   "Detalle Saldo Cuenta 41"
         End
         Begin VB.Menu mnuF0312 
            Caption         =   "Detalle Saldo Cuenta 42"
         End
         Begin VB.Menu mnuF0309_45 
            Caption         =   "Detalle Saldo Cuenta 45"
         End
         Begin VB.Menu mnuF0313 
            Caption         =   "Detalle Saldo Cuenta 46"
         End
         Begin VB.Menu mnuF0314 
            Caption         =   "Detalle Saldo Cuenta 47"
         End
         Begin VB.Menu mnuF0315 
            Caption         =   "Detalle Saldo Cuenta 49"
         End
         Begin VB.Menu mnuF0316 
            Caption         =   "Detalle Saldo Cuenta 50"
         End
         Begin VB.Menu mnuF0316_50 
            Caption         =   "Detalle Saldo Cuenta 50 (Divisionaria)"
         End
         Begin VB.Menu mnuF0316_57 
            Caption         =   "Detalle Saldo Cuenta 57"
         End
         Begin VB.Menu mnuF0316_58 
            Caption         =   "Detalle Saldo Cuenta 58"
         End
         Begin VB.Menu mnuF0316_59 
            Caption         =   "Detalle Saldo Cuenta 59"
         End
      End
      Begin VB.Menu mnuRepInvBalEmp 
         Caption         =   "Saldo de Cuentas de Balance"
         Begin VB.Menu mnuElemento01 
            Caption         =   "Elemento 1: Activo Disponible y Exigible"
            Begin VB.Menu mnuPCGE10 
               Caption         =   "10 - Efectivo y Equivalentes de Efectivo"
            End
            Begin VB.Menu mnuPCGE11 
               Caption         =   "11 - Inversiones Financieras"
            End
            Begin VB.Menu mnuPCGE12 
               Caption         =   "12 - Cuentas Por Cobrar Comerciales - Terceros"
            End
            Begin VB.Menu mnuPCGE13 
               Caption         =   "13 - Cuentas Por Cobrar Comerciales - Relacionadas"
            End
            Begin VB.Menu mnuPCGE14 
               Caption         =   "14 - Cuentas Por Cobrar Al Personal, A Los Accionistas, Directores y Gerentes"
            End
            Begin VB.Menu mnuPCGE16 
               Caption         =   "16 - Cuentas Por Cobrar diversas Terceros"
            End
            Begin VB.Menu mnuPCGE17 
               Caption         =   "17 - Cuentas Por Cobrar diversas Vinculadas"
            End
            Begin VB.Menu mnuPCGE18 
               Caption         =   "18 - Servicios y Otros Contratos  Por Anticipado"
            End
            Begin VB.Menu mnuPCGE19 
               Caption         =   "19 - Estimacion para Cuentas de Cobranza Dudosa"
            End
         End
         Begin VB.Menu mnuElemento02 
            Caption         =   "Elemento 2: Activo Realizable"
            Begin VB.Menu mnuPCGE20 
               Caption         =   "20 - Mercaderias"
            End
            Begin VB.Menu mnuPCGE21 
               Caption         =   "21 - Productos Terminados"
            End
            Begin VB.Menu mnuPCGE22 
               Caption         =   "22 - Sub Productos"
            End
            Begin VB.Menu mnuPCGE23 
               Caption         =   "23 - Productos En Proceso"
            End
            Begin VB.Menu mnuPCGE24 
               Caption         =   "24 - Materias Primas"
            End
            Begin VB.Menu mnuPCGE25 
               Caption         =   "25 - Materiales Auxiliares, Suministros y Repuestos"
            End
            Begin VB.Menu mnuPCGE26 
               Caption         =   "26 - Envases y Embalajes"
            End
            Begin VB.Menu mnuPCGE27 
               Caption         =   "27 - Activos No corrientes Mantenidos Para La Venta"
            End
            Begin VB.Menu mnuPCGE28 
               Caption         =   "28 - Existencias Por Recibir"
            End
            Begin VB.Menu mnuPCGE29 
               Caption         =   "29 - Desvalorizacion de Existencias"
            End
         End
         Begin VB.Menu mnuElemento03 
            Caption         =   "Elemento 3: Activo Inmovilizado"
            Begin VB.Menu mnuPCGE30 
               Caption         =   "30 - Inversiones Mobiliarias"
            End
            Begin VB.Menu mnuPCGE31 
               Caption         =   "31 - Inversiones Inmobiliarias"
            End
            Begin VB.Menu mnuPCGE32 
               Caption         =   "32 - Activos Adquiridos En Arrendamiento Financiero"
            End
            Begin VB.Menu mnuPCGE33 
               Caption         =   "33 - Inmuebles, Maquinarias y Equipos"
            End
            Begin VB.Menu mnuPCGE34 
               Caption         =   "34 - Intangibles"
            End
            Begin VB.Menu mnuPCGE35 
               Caption         =   "35 - Activos Biologicos"
            End
            Begin VB.Menu mnuPCGE36 
               Caption         =   "36 - Desvalorizacion De Activos Inmovilizados"
            End
            Begin VB.Menu mnuPCGE37 
               Caption         =   "37 - Activo Diferido"
            End
            Begin VB.Menu mnuPCGE38 
               Caption         =   "38 - Otros Activos"
            End
            Begin VB.Menu mnuPCGE39 
               Caption         =   "39 - Depreciacion, Amortizacion y Agotamientos Acumulados"
            End
         End
         Begin VB.Menu mnuElemento04 
            Caption         =   "Elemento 4: Pasivo"
            Begin VB.Menu mnuPCGE40 
               Caption         =   "40 - Tributos y Aportes al Sistema de Pensiones y de Salud"
            End
            Begin VB.Menu mnuPCGE41 
               Caption         =   "41 - Remuneraciones y PArticipaciones Por Pagar"
            End
            Begin VB.Menu mnuPCGE42 
               Caption         =   "42 - Cuentas Por Pagar Comerciales Terceros"
            End
            Begin VB.Menu mnuPCGE43 
               Caption         =   "43 - Cuentas Por Pagar Comerciales Relacionadas"
            End
            Begin VB.Menu mnuPCGE44 
               Caption         =   "44 - Cuentas Por Pagar A Los Accionistas, Directores y Gerentes"
            End
            Begin VB.Menu mnuPCGE45 
               Caption         =   "45 - Obligaciones Financieras"
            End
            Begin VB.Menu mnuPCGE46 
               Caption         =   "46 - Cuentas Por Pagar Diversas Terceros"
            End
            Begin VB.Menu mnuPCGE47 
               Caption         =   "47 - Cuentas Por Pagar Diversas Relacionadas"
            End
            Begin VB.Menu mnuPCGE48 
               Caption         =   "48 - Provisiones Diversas"
            End
            Begin VB.Menu mnuPCGE49 
               Caption         =   "49 - Pasivo Diferido"
            End
         End
         Begin VB.Menu mnuElemento05 
            Caption         =   "Elemento 5: Patrimonio"
            Begin VB.Menu mnuPCGE50 
               Caption         =   "50 - Capital"
               Index           =   0
            End
            Begin VB.Menu mnuPCGE50 
               Caption         =   "50 - Capital (Divisionaria)"
               Index           =   1
            End
            Begin VB.Menu mnuPCGE51 
               Caption         =   "51 - Acciones de Inversion"
            End
            Begin VB.Menu mnuPCGE52 
               Caption         =   "52 - Capital Adicional"
            End
            Begin VB.Menu mnuPCGE56 
               Caption         =   "56 - Resultados No Realizados"
            End
            Begin VB.Menu mnuPCGE57 
               Caption         =   "57 - Excedente De Revaluacion"
            End
            Begin VB.Menu mnuPCGE58 
               Caption         =   "58 - Reservas"
            End
            Begin VB.Menu mnuPCGE59 
               Caption         =   "59 - Resultados Acumulados"
            End
         End
      End
      Begin VB.Menu mnuLinea08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepCajaBancos 
         Caption         =   "Libro Caja y Bancos"
         Begin VB.Menu mnuF0101 
            Caption         =   "Detalle del movimiento efectivo"
         End
         Begin VB.Menu mnuF0102 
            Caption         =   "Detalle de cuenta corriente"
         End
      End
      Begin VB.Menu mnuF0501 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu mnuF0601 
         Caption         =   "Libro Mayor"
      End
      Begin VB.Menu mnuF0801 
         Caption         =   "Registro de Compras"
      End
      Begin VB.Menu mnuF1401 
         Caption         =   "Registro de Ventas"
      End
      Begin VB.Menu mnuLinea06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF0317 
         Caption         =   "Balance de Comprobación"
      End
      Begin VB.Menu mnuF0317_2 
         Caption         =   "Balance de Comprobación Anual - PDT"
      End
      Begin VB.Menu mnuF0301 
         Caption         =   "Estado de Situación Financiera"
      End
      Begin VB.Menu mnuF0320 
         Caption         =   "Estado del Resultado Integral - Funcion"
      End
      Begin VB.Menu mnuF0321 
         Caption         =   "Estado del Resultado Integral - Naturaleza"
      End
      Begin VB.Menu mnuF0319 
         Caption         =   "Estado de Cambios en el Patrimonio"
      End
      Begin VB.Menu mnuDiarioSimpRep 
         Caption         =   "Diario Simplificado"
      End
      Begin VB.Menu mnulinea07 
         Caption         =   "-"
      End
      Begin VB.Menu regPercion 
         Caption         =   "Registro de Percepciones y Retenciones"
      End
      Begin VB.Menu mnuReporteDaot 
         Caption         =   "Reporte de DAOT"
      End
   End
   Begin VB.Menu mnuGerencial 
      Caption         =   "&Gerencial"
      Begin VB.Menu mnuIndicadores 
         Caption         =   "Registro de Indicadores"
      End
      Begin VB.Menu mnuRatiosFinancieros 
         Caption         =   "Reporte de Ratios Financieros"
      End
   End
   Begin VB.Menu mnuUtilitarios 
      Caption         =   "&Utilitarios"
      Begin VB.Menu mnuDatos 
         Caption         =   "Datos"
         Begin VB.Menu mnuExportarDatosSistema 
            Caption         =   "Exportar Datos"
         End
         Begin VB.Menu mnuImportarDatosSistema 
            Caption         =   "Importar Datos"
         End
         Begin VB.Menu mnuVerAsientosImportados 
            Caption         =   "Ver Asientos Importados"
         End
         Begin VB.Menu mnuAuditoriaAsientos 
            Caption         =   "Auditoria de Asientos"
         End
         Begin VB.Menu mnuExportarPDB 
            Caption         =   "Exportar PDB"
         End
         Begin VB.Menu mnuExportarDAOT 
            Caption         =   "Exportar DAOT"
         End
         Begin VB.Menu mnuExportarPDT0601 
            Caption         =   "Exportar PDT 0601 - PLAME"
         End
         Begin VB.Menu mnuGuion23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportarDatos 
            Caption         =   "Importar Datos Otros Sistemas"
         End
      End
      Begin VB.Menu mnuGuion24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackupBDatos 
         Caption         =   "Backup de la Base de Datos"
      End
      Begin VB.Menu mnuRestaurarBaseDatos 
         Caption         =   "Restaurar la Base de Datos"
      End
      Begin VB.Menu mnuGuion15Utilitarios 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndexar 
         Caption         =   "Indexación de Base de Datos"
      End
      Begin VB.Menu mnuGuion25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibroElec 
         Caption         =   "Libros Electrónicos"
         Begin VB.Menu mnuGenLibroElec 
            Caption         =   "Generar Libro Electrónico"
         End
         Begin VB.Menu mnuConLibroElec 
            Caption         =   "Consulta Libro Electrónico"
         End
         Begin VB.Menu mnuUpdLibroElec 
            Caption         =   "Actualizar Estados Libros Electrónicos"
         End
      End
      Begin VB.Menu mnulin06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActualizarAsientosDestino 
         Caption         =   "Actualizar Asientos de Destino"
      End
      Begin VB.Menu mnuActualizarSaldos 
         Caption         =   "Actualizar Saldos"
      End
      Begin VB.Menu mnuActRecCentrosCosto 
         Caption         =   "Actualizar Centro de Costos"
      End
      Begin VB.Menu mnuGuion26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPeriodo 
         Caption         =   "Apertura de Periodo Contable"
      End
      Begin VB.Menu mnuGuion28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpresa 
         Caption         =   "Mantenimiento de Empresas"
      End
      Begin VB.Menu mnuGuion29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Mantenimiento de Usuarios"
      End
      Begin VB.Menu mnuEmpUsr 
         Caption         =   "Mantenimiento de Usuario / Empresa"
      End
      Begin VB.Menu mnuEmpUserLibro 
         Caption         =   "Mantenimiento de Usuario / Libro"
      End
      Begin VB.Menu mnuAccesos 
         Caption         =   "Mantenimiento de Acceso al Sistema"
      End
      Begin VB.Menu mnuPerfiles 
         Caption         =   "Mantenimiento de Perfiles del Sistema"
      End
   End
   Begin VB.Menu m_Meses 
      Caption         =   "&Meses"
      Visible         =   0   'False
      Begin VB.Menu m_mes 
         Caption         =   " APERTURA"
         Index           =   0
      End
      Begin VB.Menu m_mes 
         Caption         =   " ENERO"
         Index           =   1
      End
      Begin VB.Menu m_mes 
         Caption         =   " FEBRERO"
         Index           =   2
      End
      Begin VB.Menu m_mes 
         Caption         =   " MARZO"
         Index           =   3
      End
      Begin VB.Menu m_mes 
         Caption         =   " ABRIL"
         Index           =   4
      End
      Begin VB.Menu m_mes 
         Caption         =   " MAYO"
         Index           =   5
      End
      Begin VB.Menu m_mes 
         Caption         =   " JUNIO"
         Index           =   6
      End
      Begin VB.Menu m_mes 
         Caption         =   " JULIO"
         Index           =   7
      End
      Begin VB.Menu m_mes 
         Caption         =   " AGOSTO"
         Index           =   8
      End
      Begin VB.Menu m_mes 
         Caption         =   " SETIEMBRE"
         Index           =   9
      End
      Begin VB.Menu m_mes 
         Caption         =   " OCTUBRE"
         Index           =   10
      End
      Begin VB.Menu m_mes 
         Caption         =   " NOVIEMBRE"
         Index           =   11
      End
      Begin VB.Menu m_mes 
         Caption         =   " DICIEMBRE"
         Index           =   12
      End
      Begin VB.Menu m_mes 
         Caption         =   " AJUSTE"
         Index           =   13
      End
      Begin VB.Menu m_mes 
         Caption         =   " CIERRE"
         Index           =   14
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      Visible         =   0   'False
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Casca&da"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Mosaico &horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Mosaico &vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Organizar iconos"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuIndice 
         Caption         =   "&Indice"
      End
      Begin VB.Menu mnuBusqueda 
         Caption         =   "&Busqueda"
      End
      Begin VB.Menu mnule 
         Caption         =   "&Libros Electrónicos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnulin10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "&Acerca de EcbCont ..."
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "Version"
         Shortcut        =   +^{F12}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu mnuGrupos 
         Caption         =   "Grupos"
         Begin VB.Menu mnuGru_Agrupar 
            Caption         =   "Agruprar"
         End
         Begin VB.Menu mnuGru_Desagrupar 
            Caption         =   "Desagrupar"
         End
         Begin VB.Menu mnulin11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGru_Copiar 
            Caption         =   "Copiar Grupo"
         End
         Begin VB.Menu mnuGru_Pegar 
            Caption         =   "Pegar Grupo"
         End
         Begin VB.Menu mnulin12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGru_DesagTodo 
            Caption         =   "Desagrupar Todo"
         End
      End
      Begin VB.Menu mnuDiarioSimplificado 
         Caption         =   "Diario Simplificado"
         Begin VB.Menu mnuAsigConcepto 
            Caption         =   "Asignar Concepto de Operación"
         End
         Begin VB.Menu mnuQuitarConcepto 
            Caption         =   "Quitar Concepto de Operación"
         End
         Begin VB.Menu mnulin05 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTodosConceptos 
            Caption         =   "Quitar todos los Conceptos de Operación"
         End
      End
   End
   Begin VB.Menu mnuOpcion 
      Caption         =   "mnuOpcion"
      Visible         =   0   'False
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular Documento"
      End
      Begin VB.Menu mnuRevertir 
         Caption         =   "Revertir Anulación"
      End
   End
   Begin VB.Menu mnuCerrarTab 
      Caption         =   "CerrarTab"
      Visible         =   0   'False
      Begin VB.Menu mnuCerrarTabVentana 
         Caption         =   "Cerrar Ventana"
      End
   End
End
Attribute VB_Name = "frmMDIConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ActivaMenuEmergente As Boolean
Dim Logoff As Boolean
Dim CancelSalida As Boolean
Dim Salida As Boolean
Dim oReportesEmp As frmRepAnexoInvBalanceEmp
Dim oReportesEmpAnt As frmRepAnexoInvBalance
Dim oReportesEmpAntBcos As frmRepLibroBancos

Public Sub DesactivaMenuNoActivo(CodEmpresa As String)
    'Despues de activar o desactivar por tabla de accesos,
    'hace una segunda desactivacion por tabla principal de menu
    Dim sqlSp As String
    Dim arrDatos() As Variant
    Dim clDatos  As New clsMantoTablas
    Dim lrsTabla As New ADODB.Recordset
    
    sqlSp = "spSg_GrabaAcceso 'SEL_ALL_MENUDESAC', '" & gsSOFT & "', '" & gsUsuario & "', '','','" & _
    CodEmpresa & "'"
    arrDatos = Array(sqlSp)
    
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Not lrsTabla Is Nothing Then
        If lrsTabla.State = adStateOpen Then
            lrsTabla.MoveFirst
            Do While Not lrsTabla.EOF
                ActivaControl CE(lrsTabla!opm_cdesmenu), CE(lrsTabla!opm_cnomobj), False
                lrsTabla.MoveNext
            Loop
        End If
    End If
    
    Call CerrarRecordSet(lrsTabla)
    Set clDatos = Nothing
    Set lrsTabla = Nothing
End Sub

Public Sub CargaValoresMenu(CodEmpresa As String)
    Dim sqlSp As String
    Dim cOpcion As String
    Dim arrDatos() As Variant
    Dim clDatos  As New clsMantoTablas
    Dim lrsTabla As New ADODB.Recordset
    
    sqlSp = "spSg_GrabaAcceso 'SEL_ALL_INVERT', '" & gsSOFT & "', '" & gsUsuario & "', '','','" & CodEmpresa & "'"

    arrDatos = Array(sqlSp)
    
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    '*******************************
    ' ALMACENA LOS ACCESOS A CADA MENU EN EL ARRAY DE ACCESOS
    Call AlmacenaArray(lrsTabla)
    '*******************************
    If Not lrsTabla Is Nothing Then
        If lrsTabla.State = adStateOpen Then
            lrsTabla.MoveFirst
            Do While Not lrsTabla.EOF
                ConfigMenu CE(lrsTabla!opm_cnomobj), IIf(CE(lrsTabla!OPM_CACTIVADO) = "0", False, True), CE(lrsTabla!PFL_CCODPERFIL)
                
                If CE(lrsTabla!opm_cnomobj) = "mnuReporteAsientosEntidades" Then
                    sqlSp = sqlSp
                End If
                
                cOpcion = CE(lrsTabla!OPM_CACTIVADO)
                
                If cOpcion = "0" Then
                    ActivaControl CE(lrsTabla!opm_cdesmenu), CE(lrsTabla!opm_cnomobj), False
                Else
                    ActivaControl CE(lrsTabla!opm_cdesmenu), CE(lrsTabla!opm_cnomobj), True
                End If

                lrsTabla.MoveNext
            Loop
        End If
    End If
    Call CerrarRecordSet(lrsTabla)
    Set clDatos = Nothing
    Set lrsTabla = Nothing
End Sub

Private Sub ConfigMenu(menu As String, Visible As Boolean, sPrivilegio As String)
On Error Resume Next
    Me.Controls(menu).Visible = IIf(Me.Controls(menu).Check = True, False, Visible)
    If Me.Controls(menu).Check = False Then Me.Controls(menu).Enabled = Visible
    '*** INVERTIR EL PRIVILEGIO SI ES ADMINISTRADOR
    Me.Controls(menu).Tag = GenMenuUserProfile(IIf(sPrivilegio = "0000", "1111", sPrivilegio))
    Exit Sub
End Sub

Private Function GenMenuUserProfile(ByVal sPrivilegio As String) As String
Dim sProfile As String
    sPrivilegio = Left(Trim(sPrivilegio) & "0000", 4)
    sProfile = Mid(sPrivilegio, 2, 1) _
            & "1" _
            & "1" _
            & Mid(sPrivilegio, 4, 1) _
            & Mid(sPrivilegio, 3, 1) _
            & Mid(sPrivilegio, 1, 1) _
            & "1" _
            & "1"
    '*** PROFILE = NUEVO + CONSULTA + GRABAR + ELIMINAR + MODIFICAR + ELIMINAR + IMPRIMIR + CANCELAR + SALIR
    GenMenuUserProfile = sProfile
End Function

Private Sub ActivaControl(Descripcion As String, Control As String, Valor As Boolean)

On Error GoTo ERROR
    If frmMDIConta.Controls(Control).Enabled = True Then
        frmMDIConta.Controls(Control).Visible = Valor
    Else
        frmMDIConta.Controls(Control).Visible = False
    End If
    Exit Sub
ERROR:
    
'    Mensajes "Control : " & Control & Chr(10) + Chr(13) & "Descripción: " & Descripcion & Chr(10) + Chr(13) & "Error: " & Err.Description, vbOKOnly + vbInformation
End Sub

Private Sub MDIForm_Activate()
     Call ConfigurarBarraEstado
End Sub

Private Sub MDIForm_Load()
    Salida = False
    Logoff = False
    Me.Top = 0
    
    If SW_ActPLE = "1" Then
        frmMDIConta.mnuGuion25.Visible = True
        frmMDIConta.mnuLibroElec.Visible = True
    Else
        frmMDIConta.mnuGuion25.Visible = False
        frmMDIConta.mnuLibroElec.Visible = False
    End If
    
    Call OpcionesUsuario
    Call CargaVariablesMonedas
    '------------------------------------
    Call CargaValoresMenu(gsEmpresa) 'LEE TABLA ACCESOS Y ALMACENA LOS ACCESOS A CADA MENU EN EL ARRAY DE ACCESOS
    Call DesactivaMenuNoActivo(gsEmpresa) 'LEE LOS DESACTIVADOS DE TABLA MENU Y NO LOS MUESTRA EN EL SISTEMA
    '------------------------------------
    
    Call BuscaCuentas6o9
    Call pCargaCfgCtas
    
    'Valida si PLE esta activado o no
    SW_ActPLE = IIf(IsNull(ExtraeCampo("ActPLE", "SGM_SOFTWARE", "sof_cCodSoft = '001'")) = True, "0", ExtraeCampo("ActPLE", "SGM_SOFTWARE", "sof_cCodSoft = '001'"))
    
    If SW_ActPLE = "1" Then
        frmMDIConta.mnuGuion25.Visible = True
        frmMDIConta.mnuLibroElec.Visible = True
    Else
        frmMDIConta.mnuGuion25.Visible = False
        frmMDIConta.mnuLibroElec.Visible = False
    End If
    
    If gsPLE = "0" Then
        mnuLibroElec.Enabled = False
    Else
        mnuLibroElec.Enabled = True
    End If
    
        Dim Resolucion As Double, imagen As String
        With Screen
            Resolucion = (.Width \ .TwipsPerPixelX)
        End With
                
        If Resolucion < 768 Then
            Set Me.Picture = picFondo2.Picture
        Else
             Set Me.Picture = picFondo1.Picture
        End If
    
    If pCargaCfgLibro = False Then frmManSubDiarioTDoc.Show
    
    'Call ActivaMenuSegunTipoPlan
    
    nContadorProc = 0
    '*** LIMPIAR TAB STRIP
    tsForms.Tabs.Clear
End Sub

Public Sub ActivaMenuSegunTipoPlan()
    
    'If gsAMBOSPLANDECUENTAS = True Then Exit Sub
    
    If gsTipoPlan = 0 Then 'PLAN DE CUENTAS REVISADO
        mnuRepInvBal.Visible = True
        mnuRepInvBalEmp.Visible = False
    Else
        mnuRepInvBal.Visible = False
        mnuRepInvBalEmp.Visible = True
    End If
    
End Sub

Private Sub ActivarMenuControl(oObjeto As Object, nvalor As Boolean)
    oObjeto.Enabled = nvalor
    oObjeto.Visible = nvalor
End Sub

Private Function BuscaCuentas6o9() As Boolean
    Dim sqlCta As String
    
    sqlCta = "SELECT count(*) from CNA_CTAS_CONDESTINO WHERE Emp_cCodigo = '" & gsEmpresa & "' AND " & _
             "Cde_cEstado = 'A' "
    If ExisteDato(sqlCta) = False Then
On Error GoTo serror
        Dim cn As ADODB.Connection
        Set cn = New ADODB.Connection
        cn.ConnectionString = gsCadenaConexion
        cn.Open
        cn.Execute "INSERT CNA_CTAS_CONDESTINO(Emp_cCodigo, Cde_cClase, Cde_cEstado, Cde_cUserCrea, Cde_dFechaCrea, Cde_cUserModifica, Cde_dFechaModifica, Cde_cEquipoUser) " & _
                   "VALUES ('" & gsEmpresa & "','6','A','" & gsUsuario & "','" & Date & "','" & gsUsuario & "','" & Date & "','')"
    
        cn.Execute "INSERT CNA_CTAS_CONDESTINO(Emp_cCodigo, Cde_cClase, Cde_cEstado, Cde_cUserCrea, Cde_dFechaCrea, Cde_cUserModifica, Cde_dFechaModifica, Cde_cEquipoUser) " & _
                   "VALUES ('" & gsEmpresa & "','9','A','" & gsUsuario & "','" & Date & "','" & gsUsuario & "','" & Date & "','')"
serror:
        
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
        
    End If
End Function

Private Sub CargaVariablesMonedas()
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCNT_TIPO_MONEDA  'BUSCARACTIVOS','" & gsEmpresa & "','','','','','','','','',''"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        MsgBox "Identifique el tipo de moneda Nacional o Extranjera de sistema", vbInformation + vbOKOnly, "Cuidado..."
        frmManTipoMoneda.Show
    Else
        Do While Not rsArreglo.EOF
            If rsArreglo.AbsolutePosition = 1 Then
                gsMonedaNac = CE(rsArreglo("codigo"))
                gsNombreMonedaNac = CE(rsArreglo("descripcion"))
                gsMonedaNacAbrev = CE(rsArreglo("Abrev"))
            End If
            If rsArreglo.AbsolutePosition = 2 Then
                gsMonedaExt = CE(rsArreglo("codigo"))
                gsNombreMonedaExt = CE(rsArreglo("descripcion"))
                gsMonedaExtAbrev = CE(rsArreglo("Abrev"))
            End If
            rsArreglo.MoveNext
        Loop
    End If
    CerrarRecordSet rsArreglo
End Sub

Private Sub OpcionesUsuario()
    ' *** En esta opcion se habilitara o deshabilitara las opciones del usuario
    Select Case GrupoUsuario(gsUsuario)
            Case "002"
                 mnuUsuarios.Enabled = False
                 mnuEmpUsr.Enabled = False
                 mnuEmpresa.Enabled = False
                 mnuPeriodo.Enabled = False
                 mnuAuditoriaAsientos.Enabled = False
            Case "003"
                 mnuPlanCuentas.Enabled = False
                 mnuCentroCostos.Enabled = False
                 mnuTablasMnt.Enabled = False
                 mnuProcesos.Enabled = False
                 mnuUtilitarios.Enabled = False
                 mnuAuditoriaAsientos.Enabled = False
            Case "004"
                 mnuTablas.Enabled = False
                 mnuIngresos.Enabled = False
                 mnuProcesos.Enabled = False
                 mnuConciliacionBancaria.Enabled = False
                 mnuGerencial.Enabled = False
                 mnuUtilitarios.Enabled = False
                 tbrOpciones.Buttons(1).Enabled = False
                 tbrOpciones.Buttons(2).Enabled = False
                 tbrOpciones.Buttons(3).Enabled = False
                 mnuAuditoriaAsientos.Enabled = False
    End Select

End Sub

Private Function CerrarVentanas(Optional ByVal sPropCambio As String = "", Optional bPreguntar As Boolean = True) As Boolean

Dim sMsg As String
CerrarVentanas = False
    If Me.ActiveForm Is Nothing Then CerrarVentanas = True: Exit Function
    sMsg = "Esta operación cerrara TODAS la ventanas activas"
    sMsg = sMsg & Chr(13) & IIf(sPropCambio <> "", "¿Desea Cambiar de " & sPropCambio & " de todas maneras?", "")
    
    If bPreguntar Then
        If MensajesRet(sMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Do While Not Me.ActiveForm Is Nothing
                Unload Me.ActiveForm
            Loop
            CerrarVentanas = True
        End If
    Else
        
        Do While Not Me.ActiveForm Is Nothing
            Unload Me.ActiveForm
        Loop
        
        CerrarVentanas = True
    End If
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If CerrarVentanas = False Then Cancel = 1
    End If
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        If Me.Width < 11880 Then
            Me.Width = 11880
        End If
        If Me.Height < 8220 Then
            Me.Height = 8220
        End If
    End If
    
    If Me.WindowState <> vbMinimized Then
        Call TabForms_Redim
    End If

End Sub
Public Function SalirSistema() As Boolean
    MDIForm_Unload (0)
    SalirSistema = CancelSalida
End Function
Private Sub MDIForm_Unload(Cancel As Integer)

    Salida = False
    
    Dim strMensaje As String
    CancelSalida = False
    
    If Logoff = True Then
        strMensaje = "Desea salir del sistema e ingresar con otro usuario"
    Else
        strMensaje = "Desea salir del sistema"
    End If
    
    If MsgBox(strMensaje, vbYesNo + vbQuestion, "Aviso...") = vbYes Then
        
        If Logoff = True Then
            gsInicio = False
            
                frmPrcIngresoSistema.Show
            
            Cancel = 0
            CancelSalida = False
        Else
            Cancel = 0
            CancelSalida = False

            Set frmMDIConta = Nothing
            End
        End If
    Else
        CancelSalida = True
        Cancel = 1
        Logoff = False
    End If
    
End Sub

Private Sub mnu_reg_aux_ventas_Click()
    'Call GenerarVentanaAntigua("F0303", mnuF0303.Caption, "12")
    FrmManRegAuxiliarVentas.Grupo = BuscaArray("mnu_reg_aux_ventas")
    FrmManRegAuxiliarVentas.Show
    pSetFocus FrmManRegAuxiliarVentas
End Sub

Private Sub mnuAccesos_Click()
    frmManMenu.Grupo = BuscaArray("mnuAccesos")
    frmManMenu.Show
    pSetFocus frmManMenu
End Sub

Private Sub mnuActRecCentrosCosto_Click()
 frmPrcCentroCosto.Grupo = BuscaArray("mnuActRecCentrosCosto")
 frmPrcCentroCosto.Show
 pSetFocus frmPrcCentroCosto
End Sub

Private Sub mnuActualizarAsientosDestino_Click()
 frmPrcActualizaDestino.Grupo = BuscaArray("mnuActualizarAsientosDestino")
 frmPrcActualizaDestino.Show
 pSetFocus frmPrcActualizaDestino
End Sub

Private Sub mnuActualizarSaldos_Click()
    frmPrcActualizaSaldos.Grupo = BuscaArray("mnuActualizarSaldos")
    frmPrcActualizaSaldos.Show
    pSetFocus frmPrcActualizaSaldos
End Sub

Private Sub mnuArrendaFinanciero_Click()
'    FrmManRegActRenFin.Show
'    pSetFocus FrmManRegActRenFin
'    Exit Sub
    FrmManRegActRenFin.Grupo = BuscaArray("mnuArrendaFinanciero")
    FrmManRegActRenFin.Show
End Sub

Private Sub mnuAsientoApertura_Click()
    frmPrcAsientoApertura.Grupo = BuscaArray("mnuAsientoApertura")
    frmPrcAsientoApertura.Show
    pSetFocus frmPrcAsientoApertura
End Sub

Private Sub mnuAsientoCierre_Click()
    frmPrcAsientoCierre.Grupo = BuscaArray("mnuAsientoCierre")
    frmPrcAsientoCierre.Show
    pSetFocus frmPrcAsientoCierre
End Sub

Private Sub mnuAsientosCentroCosto_Click()
    frmRepAsientosCCosto.Grupo = BuscaArray("mnuAsientosCentroCosto")
    frmRepAsientosCCosto.Show
    pSetFocus frmRepAsientosCCosto
End Sub

Private Sub mnuAuditoriaAsientos_Click()
    frmConAuditoriaAsientos.Grupo = BuscaArray("mnuAuditoriaAsientos")
    frmConAuditoriaAsientos.Show
    pSetFocus frmConAuditoriaAsientos
End Sub

Private Sub mnuBackupBDatos_Click()
    frmPrcBackup.Grupo = BuscaArray("mnuBackupBDatos")
    frmPrcBackup.Show
    pSetFocus frmPrcBackup
End Sub

Private Sub mnuBalanceComprobacion_Click()
    If BuscaForm("frmRepBalanceComprobacion") Then frmRepBalanceComprobacion.CerrarForm
    
    frmRepBalanceComprobacion.Grupo = BuscaArray("mnuBalanceComprobacion")
    frmRepBalanceComprobacion.Show
    pSetFocus frmRepBalanceComprobacion
End Sub

Private Sub MnuBalGral_Click()
FrmBalanceGeneralCompMensual.Show
End Sub

Private Sub mnuBancos_Click()
    frmManBancos.Grupo = BuscaArray("mnuBancos")
    frmManBancos.Show
    pSetFocus frmManBancos

End Sub

Private Sub mnuCambioAnio_Click()
    If CerrarVentanas("") = True Then
            gsCambioEmpresa = True
            gsInicio = False
            frmBusPeriodo.Grupo = BuscaArray("mnuCambioAnio")
            frmBusPeriodo.Show
            frmBusPeriodo.Caption = " Cambio de Año de Trabajo"
            pSetFocus frmBusPeriodo
'            gsTipoPlan = ExtraeCampo("Cfl_cTipoPlan", "CNT_CONFIG_LIBROS", "Emp_cCodigo='" & gsEmpresa & "' And Pan_cAnio='" & gsAnio & "'")
'            '----------------------------------------
'            Call frmMDIConta.ActivaMenuSegunTipoPlan
'            '----------------------------------------
    End If
'    gsInicio = True
End Sub

Private Sub mnuCapital_Click()
    frmManCapital.Grupo = BuscaArray("mnuCapital")
    frmManCapital.Show
    pSetFocus frmManCapital

End Sub

Private Sub mnuCentroCostos_Click()
    frmManCentroCostoNiv.Grupo = BuscaArray("mnuCentroCostos")
    frmManCentroCostoNiv.Show
    pSetFocus frmManCentroCostoNiv
    
End Sub

Private Sub mnuChequesPendientes_Click()
    frmRepChequesPendientes.Grupo = BuscaArray("mnuChequesPendientes")
    frmRepChequesPendientes.Show
    pSetFocus frmRepChequesPendientes
End Sub

Private Sub mnuCierreEjercicio_Click()
    frmPrcCierreEjercicio.Grupo = BuscaArray("mnuCierreMensual")
    frmPrcCierreEjercicio.Show
    pSetFocus frmPrcCierreEjercicio
End Sub

Private Sub mnuCierreMensual_Click()
    frmPrcCierreMes.Grupo = BuscaArray("mnuCierreMensual")
    frmPrcCierreMes.Show
    pSetFocus frmPrcCierreMes
End Sub

Private Sub mnuConfigOperac_Click()
    frmConfigOperaciones.Grupo = BuscaArray("mnuConfigOperac")
    frmConfigOperaciones.Show
    pSetFocus frmConfigOperaciones
    
End Sub

Private Sub mnuConLibroElec_Click()
    frmConsultarLE.Show
End Sub

Private Sub mnuCuentaCorriente_Click()
    frmManCuentaCorriente.Grupo = BuscaArray("mnuCuentaCorriente")
    frmManCuentaCorriente.Show
    pSetFocus frmManCuentaCorriente
End Sub

Private Sub mnuDepA_Click()
    frmManDepAduanera.Show
    
End Sub

Private Sub mnuDiarioSimpRep_Click()
    Call GenerarVentanaAntigua("F0502", "Formato 5.2 : " & mnuDiarioSimpRep.Caption, "")
End Sub

Private Sub MnuEGPPF_Click()
    FrmEstadoGananciasPerdidasXFunc.Show
End Sub

Private Sub mnuEmpresa_Click()
    frmManEmpresas.Grupo = BuscaArray("mnuEmpresa")
    frmManEmpresas.Show
    pSetFocus frmManEmpresas
End Sub

Private Sub mnuEmpUserLibro_Click()
    frmManUsrEmpLib.Grupo = BuscaArray("mnuEmpUserLibro")
    frmManUsrEmpLib.Show
    pSetFocus frmManUsrEmpLib
End Sub

Private Sub mnuEmpUsr_Click()
    frmManUsrEmp.Grupo = BuscaArray("mnuEmpUsr")
    frmManUsrEmp.Show
    pSetFocus frmManUsrEmp
End Sub

Private Sub mnuEntDoc_Click()
    frmManCfgEntDoc.Grupo = BuscaArray("mnuEntDoc")
    frmManCfgEntDoc.Show
    pSetFocus frmManCfgEntDoc
    
End Sub

Private Sub mnuEntidadesMnt_Click()
    frmManEntidades.Grupo = BuscaArray("mnuEntidadesMnt")
    frmManEntidades.Show
    pSetFocus frmManEntidades
    
End Sub

Private Sub mnuExportarPDB_Click()
    frmPrcExportaPDB.Grupo = BuscaArray("mnuExportarPDB")
    frmPrcExportaPDB.Show
    pSetFocus frmPrcExportaPDB
End Sub

Private Sub mnuExportarPDT0601_Click()
    frmPrcExportarPDT.Grupo = BuscaArray("mnuExportarPDT0601")
    frmPrcExportarPDT.Show
    pSetFocus frmPrcExportarPDT
End Sub

Private Sub mnuF0317_2_Click()
    If BuscaForm("frmRepBalanceComprobacion") Then frmRepBalanceComprobacion.CerrarForm
    frmRepBalanceComprobacion.Grupo = BuscaArray("mnuBalanceComprobacion")
    frmRepBalanceComprobacion.TituloSunat = "Formato 3.17 : " & mnuF0317_2.Caption
    frmRepBalanceComprobacion.Show
    frmRepBalanceComprobacion.ReporteSunat = "F0317_2"
    frmRepBalanceComprobacion.optDelMes.Visible = False
    pSetFocus frmRepBalanceComprobacion
End Sub

Private Sub mnuGenLibroElec_Click()
frmLibrosElectronicos.Show
End Sub

Private Sub mnuImportarDatos2_Click()
    Form1.Grupo = BuscaArray("mnuImportarDatos")
    Form1.Show
    pSetFocus Form1
End Sub

Private Sub mnule_Click()
ShowSearch 2
End Sub

Private Sub mnuPCGE40_Click()
 Call GenerarVentana("PCGE40", mnuPCGE40.Caption)
End Sub

Private Sub mnuPCGE41_Click()
 Call GenerarVentana("PCGE41", mnuPCGE41.Caption)
End Sub

Private Sub mnuPCGE42_Click()
 Call GenerarVentana("PCGE42", mnuPCGE42.Caption)
End Sub

Private Sub mnuPCGE43_Click()
 Call GenerarVentana("PCGE43", mnuPCGE43.Caption)
End Sub

Private Sub mnuPCGE44_Click()
 Call GenerarVentana("PCGE44", mnuPCGE44.Caption)
End Sub

Private Sub mnuPCGE45_Click()
 Call GenerarVentana("PCGE45", mnuPCGE45.Caption)
End Sub

Private Sub mnuPCGE46_Click()
 Call GenerarVentana("PCGE46", mnuPCGE46.Caption)
End Sub

Private Sub mnuPCGE47_Click()
 Call GenerarVentana("PCGE47", mnuPCGE47.Caption)
End Sub

Private Sub mnuPCGE48_Click()
 Call GenerarVentana("PCGE48", mnuPCGE48.Caption)
End Sub

Private Sub mnuPCGE49_Click()
 Call GenerarVentana("PCGE49", mnuPCGE49.Caption)
End Sub

Private Sub mnuPCGE50_Click(Index As Integer)
    Call GenerarVentana("PCGE50_" & CStr(Index), mnuPCGE50(Index).Caption)
End Sub

Private Sub mnuPCGE51_Click()
 Call GenerarVentana("PCGE51", mnuPCGE51.Caption)
End Sub

Private Sub mnuPCGE52_Click()
 Call GenerarVentana("PCGE52", mnuPCGE52.Caption)
End Sub

Private Sub mnuPCGE56_Click()
 Call GenerarVentana("PCGE56", mnuPCGE56.Caption)
End Sub

Private Sub mnuPCGE57_Click()
 Call GenerarVentana("PCGE57", mnuPCGE57.Caption)
End Sub

Private Sub mnuPCGE58_Click()
 Call GenerarVentana("PCGE58", mnuPCGE58.Caption)
End Sub

'Private Sub mnuPCGE50_Click()
' Call GenerarVentana("PCGE50", mnuPCGE50.Caption)
'End Sub

Private Sub mnuPCGE59_Click()
 Call GenerarVentana("PCGE59", mnuPCGE59.Caption)
End Sub

Private Sub mnuPDBFPago_Click()
    frmManPDBPagos.Grupo = BuscaArray("mnuPDBFPago")
    frmManPDBPagos.Show
    pSetFocus frmManPDBPagos
End Sub

Private Sub mnuPDBVentas_Click()
    frmManPDBVentas.Grupo = BuscaArray("mnuPDBVentas")
    frmManPDBVentas.Show
    pSetFocus frmManPDBVentas
End Sub

Private Sub mnuExportarDAOT_Click()
    frmPrcExportarDaot.Grupo = BuscaArray("mnuExportarDAOT")
    frmPrcExportarDaot.Show
    pSetFocus frmPrcExportarDaot
End Sub

Private Sub mnuExportarDatosSistema_Click()
    frmPrcExportarDatos.Grupo = BuscaArray("mnuExportarDatosSistema")
    frmPrcExportarDatos.Show
    pSetFocus frmPrcExportarDatos
End Sub

Private Sub mnuF0101_Click()
    Call GenerarVentanaAntiguaBancos("F0101", "Formato 1.1 : " & mnuF0101.Caption, "")

'    If BuscaForm("frmRepLibroBancos") Then frmRepLibroBancos.CerrarForm
'
'    frmRepLibroBancos.Grupo = BuscaArray("mnuLibroCajaBancos")
'    frmRepLibroBancos.TituloSunat = "Formato 1.1 : " & mnuF0101.Caption
'    frmRepLibroBancos.Show
'    frmRepLibroBancos.ReporteSunat = "F0101"
'
'    pSetFocus frmRepLibroBancos
End Sub

Private Sub mnuF0102_Click()
    Call GenerarVentanaAntiguaBancos("F0102", "Formato 1.2 : " & mnuF0102.Caption, "")
    
'    If BuscaForm("frmRepLibroBancos") Then frmRepLibroBancos.CerrarForm
'
'    frmRepLibroBancos.Grupo = BuscaArray("mnuLibroCajaBancos")
'    frmRepLibroBancos.TituloSunat = "Formato 1.2 : " & mnuF0102.Caption
'    frmRepLibroBancos.Show
'    frmRepLibroBancos.ReporteSunat = "F0102"
'
'    pSetFocus frmRepLibroBancos
End Sub

Private Sub mnuF0301_Click()
    Call GenerarVentanaAntigua("F0301", "Formato 3.1 : " & mnuF0301.Caption, "")
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.1 : " & mnuF0301.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0301"
'    frmRepAnexoInvBalance.chkAnexos.Visible = True
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0302_Click()
    Call GenerarVentanaAntigua("F0302", "Formato 3.2 : " & mnuF0302.Caption, "10")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.2 : " & mnuF0302.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0302"
'    frmRepAnexoInvBalance.CuentaInvBal = "10"
'    pSetFocus frmRepAnexoInvBalance
End Sub


Private Sub mnuF0303_Click()
    Call GenerarVentanaAntigua("F0303", "Formato 3.3 : " & mnuF0303.Caption, "12")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.3 : " & mnuF0303.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0303"
'    frmRepAnexoInvBalance.CuentaInvBal = "12"
'    pSetFocus frmRepAnexoInvBalance

End Sub

Private Sub mnuF0304_Click()
    Call GenerarVentanaAntigua("F0304", "Formato 3.4 : " & mnuF0304.Caption, "14")

'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.4 : " & mnuF0304.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0304"
'    frmRepAnexoInvBalance.CuentaInvBal = "14"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0305_Click()
    Call GenerarVentanaAntigua("F0305", "Formato 3.5 : " & mnuF0305.Caption, "16")


'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.5 : " & mnuF0305.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0305"
'    frmRepAnexoInvBalance.CuentaInvBal = "16"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0306_Click()
    Call GenerarVentanaAntigua("F0306", "Formato 3.6 : " & mnuF0306.Caption, "19")

'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.6 : " & mnuF0306.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0306"
'    frmRepAnexoInvBalance.CuentaInvBal = "19"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0307_2021_Click()
    Call GenerarVentanaAntigua("mnuF0307_2021", "Formato 3.7 : " & mnuF0307_2021.Caption, "2")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.7 : " & mnuF0307_2021.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "mnuF0307_2021"
'    frmRepAnexoInvBalance.CuentaInvBal = "2"
'    frmRepAnexoInvBalance.tdbcMoneda.Visible = False
'    frmRepAnexoInvBalance.Label3(1).Visible = False
'    pSetFocus frmRepAnexoInvBalance

End Sub

Private Sub mnuF0308_Click()
    Call GenerarVentanaAntigua("F0031", "Formato 3.8 : " & mnuF0308.Caption, "31")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.8 : " & mnuF0308.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0308"
'    frmRepAnexoInvBalance.CuentaInvBal = "31"
'    pSetFocus frmRepAnexoInvBalance

End Sub

Private Sub mnuF0310_Click()
    Call GenerarVentanaAntigua("F0310", "Formato 3.10 : " & mnuF0310.Caption, "40")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.10 : " & mnuF0310.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0310"
'    frmRepAnexoInvBalance.CuentaInvBal = "40"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0311_Click()
    Call GenerarVentanaAntigua("F0311", "Formato 3.11 : " & mnuF0311.Caption, "41")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.11 : " & mnuF0311.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0311"
'    frmRepAnexoInvBalance.CuentaInvBal = "41"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0312_Click()
    Call GenerarVentanaAntigua("F0312", "Formato 3.12 : " & mnuF0312.Caption, "42")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.12 : " & mnuF0312.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0312"
'    frmRepAnexoInvBalance.CuentaInvBal = "42"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0313_Click()
 Call GenerarVentanaAntigua("F0313", "Formato 3.13 : " & mnuF0313.Caption, "46")
End Sub

Private Sub mnuF0314_Click()
    Call GenerarVentanaAntigua("F0314", "Formato 3.14 : " & mnuF0314.Caption, "47")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.14 : " & mnuF0314.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0314"
'    frmRepAnexoInvBalance.CuentaInvBal = "47"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0315_Click()
    Call GenerarVentanaAntigua("F0315", "Formato 3.15 : " & mnuF0315.Caption, "49")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.15 : " & mnuF0315.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0315"
'    frmRepAnexoInvBalance.CuentaInvBal = "49"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0316_50_Click()
    Call GenerarVentanaAntigua("F0350", "Formato : " & mnuF0316_50.Caption, "50")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0316_50.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0350"
'    frmRepAnexoInvBalance.CuentaInvBal = "50"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0316_57_Click()
    Call GenerarVentanaAntigua("F0357", "Formato : " & mnuF0316_57.Caption, "57")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0316_57.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0357"
'    frmRepAnexoInvBalance.CuentaInvBal = "57"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0316_58_Click()
    Call GenerarVentanaAntigua("F0358", "Formato : " & mnuF0316_58.Caption, "58")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0316_58.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0358"
'    frmRepAnexoInvBalance.CuentaInvBal = "58"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0316_59_Click()
    Call GenerarVentanaAntigua("F0359", "Formato : " & mnuF0316_59.Caption, "59")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0316_59.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0359"
'    frmRepAnexoInvBalance.CuentaInvBal = "59"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0316_Click()
    Call GenerarVentanaAntigua("F0316", "Formato 3.16 : " & mnuF0316.Caption, "50")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.16 : " & mnuF0316.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0316"
'    frmRepAnexoInvBalance.CuentaInvBal = "50"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0317_Click()
    If BuscaForm("frmRepBalanceComprobacion") Then frmRepBalanceComprobacion.CerrarForm

    frmRepBalanceComprobacion.Grupo = BuscaArray("mnuBalanceComprobacion")
    frmRepBalanceComprobacion.TituloSunat = "Formato 3.17 : " & mnuF0317.Caption
    frmRepBalanceComprobacion.Show
    frmRepBalanceComprobacion.ReporteSunat = "F0317"
    frmRepBalanceComprobacion.optDelMes.Visible = False
    pSetFocus frmRepBalanceComprobacion
End Sub

Private Sub mnuF0318_Click()
    frmPrcFlujos.Caption = "Formato 3.18 : " '& mnuF0318.Caption
    frmPrcFlujos.Show
    pSetFocus frmPrcFlujos
End Sub

Private Sub mnuF0319_Click()
    frmRepPatrimonioNeto.Caption = "Formato 3.19 : " & mnuF0319.Caption
    frmRepPatrimonioNeto.Show
    pSetFocus frmRepPatrimonioNeto
End Sub

Private Sub mnuF0320_Click()
    Call GenerarVentanaAntigua("F0320", "Formato 3.20 : " & mnuF0320.Caption, "")

'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 3.20 : " & mnuF0320.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0320"
'    frmRepAnexoInvBalance.chkAnexos.Visible = True
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0321_Click()
    Call GenerarVentanaAntigua("F0321", "Formato 3.21 : " & mnuF0321.Caption, "")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = mnuF0321.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0321"
'    frmRepAnexoInvBalance.chkAnexos.Visible = True
'    pSetFocus frmRepAnexoInvBalance

End Sub

Private Sub mnuF0401_Click()
    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
    
    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
    frmRepAnexoInvBalance.TituloSunat = "Formato 4.1 : Libro Retenciones"
    frmRepAnexoInvBalance.Label3(1).Caption = "Moneda"
    frmRepAnexoInvBalance.Show
    frmRepAnexoInvBalance.ReporteSunat = "F0401"
    frmRepAnexoInvBalance.lblRetenciones.Visible = True
    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0501_Click()
    If BuscaForm("frmRepLibroDiario") Then frmRepLibroDiario.CerrarForm
    
    frmRepLibroDiario.Grupo = BuscaArray("mnuDiarioGeneral")
    frmRepLibroDiario.TituloSunat = "Formato 5.1 : " & mnuF0501.Caption
    frmRepLibroDiario.Show
    frmRepLibroDiario.ReporteSunat = "F0501"
    pSetFocus frmRepLibroDiario
End Sub

Private Sub mnuF0601_Click()
    If BuscaForm("frmRepLibroMayorAnalitico") Then frmRepLibroMayorAnalitico.CerrarForm
    
    frmRepLibroMayorAnalitico.Grupo = BuscaArray("mnuMayorGeneral")
    frmRepLibroMayorAnalitico.TituloSunat = "Formato 6.1 : " & mnuF0601.Caption
    frmRepLibroMayorAnalitico.Show
    frmRepLibroMayorAnalitico.ReporteSunat = "F0601"
    pSetFocus frmRepLibroMayorAnalitico

End Sub

Private Sub mnuF0801_Click()
    Call GenerarVentanaAntigua("F0801", "Formato 8.1 : " & mnuF0801.Caption, "")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 8.1 : " & mnuF0801.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0801"
'    pSetFocus frmRepAnexoInvBalance
End Sub

'Private Sub mnuF1002_Click()
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 10.2 : " & mnuF1002.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F1002"
'    pSetFocus frmRepAnexoInvBalance
'
'End Sub

'Private Sub mnuF1003_Click()
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 10.3 : " & mnuF1003.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F1003"
'    pSetFocus frmRepAnexoInvBalance
'
'End Sub

Private Sub mnuF1401_Click()
    Call GenerarVentanaAntigua("F1401", "Formato 14.1 : " & mnuF1401.Caption, "")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato 14.1 : " & mnuF1401.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F1401"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuFlujoCuentas_Click()
    frmManFlujosCuentas.Grupo = BuscaArray("mnuFlujoEjectivo")
    frmManFlujosCuentas.Show
    pSetFocus frmManFlujosCuentas
End Sub

Private Sub mnuFlujoReporte_Click()
    frmManFlujoReporte.Grupo = BuscaArray("mnuFlujoEjectivo")
    frmManFlujoReporte.Show
    pSetFocus frmManFlujoReporte
End Sub

Private Sub mnuFlujoProceso_Click()
    frmManFlujoProceso.Grupo = BuscaArray("mnuFlujoEjectivo")
    frmManFlujoProceso.Show
    pSetFocus frmManFlujoProceso
End Sub

Private Sub mnuFlujoSaldos_Click()
    frmManFlujoSaldos.Grupo = BuscaArray("mnuFlujoEjectivo")
    frmManFlujoSaldos.Show
    pSetFocus frmManFlujoSaldos
End Sub

Private Sub mnuGenerarDiferenciaCambioMensual_Click()
    frmPrcCambioCierre.Grupo = BuscaArray("mnuGenerarDiferenciaCambioMensual")
    frmPrcCambioCierre.Show
    pSetFocus frmPrcCambioCierre
End Sub

Private Sub mnuPDBCompras_Click()
    frmManPDBCompras.Grupo = BuscaArray("mnuPDBCompras")
    frmManPDBCompras.Show
    pSetFocus frmManPDBCompras
End Sub

Private Sub mnuImportarDatos_Click()
    frmPrcImportarDatosXLS.Grupo = BuscaArray("mnuImportarDatos")
    frmPrcImportarDatosXLS.Show
    pSetFocus frmPrcImportarDatosXLS
End Sub

Private Sub mnuImportarDatosSistema_Click()
    frmPrcImportarDatosSistema.Grupo = BuscaArray("mnuImportarDatosSistema")
    frmPrcImportarDatosSistema.Show
    pSetFocus frmPrcImportarDatosSistema
End Sub

Private Sub mnuIndexar_Click()
    frmPrcReindex.Grupo = BuscaArray("mnuIndexar")
    frmPrcReindex.Show
    pSetFocus frmPrcReindex
End Sub

Private Sub mnuIndicadores_Click()
    frmManIndicadores.Grupo = BuscaArray("mnuIndicadores")
    frmManIndicadores.Show
    pSetFocus frmManIndicadores
End Sub

Private Sub mnuConfigCostos_Click()
    frmManCostos.Grupo = BuscaArray("mnuCostosProd")
    frmManCostos.Show
    pSetFocus frmManCostos
End Sub

Private Sub mnuIntangibles_Click()
    frmManIntangibles.Grupo = BuscaArray("mnuIntangibles")
    frmManIntangibles.Show
    pSetFocus frmManIntangibles
    
End Sub

Private Sub mnuInvProceso_Click()
    frmManInvProc.Grupo = BuscaArray("mnuCostosProd")
    frmManInvProc.Show
    pSetFocus frmManInvProc
End Sub

Private Sub mnuMercaderias_Click()
    frmManMercaderias.Grupo = BuscaArray("mnuMercaderias")
    frmManMercaderias.Show
    pSetFocus frmManMercaderias
    
End Sub

Private Sub mnuMovimientosBancos_Click()
    frmRepMovimientosBancos.Grupo = BuscaArray("mnuMovimientosBancos")
    frmRepMovimientosBancos.Show
    pSetFocus frmRepMovimientosBancos
End Sub

Private Sub mnuParamIniciales_Click()
    frmManSubDiarioTDoc.Grupo = BuscaArray("mnuParamIniciales")
    frmManSubDiarioTDoc.Show
    pSetFocus frmManSubDiarioTDoc
End Sub

Private Sub mnuPatrimonio_Click()
    frmManPatrimonioNeto.Grupo = BuscaArray("mnuPatrimonio")
    frmManPatrimonioNeto.Show
    pSetFocus frmManPatrimonioNeto

End Sub

Private Sub mnuPDT0601_Click()
    frmManPDT0601.Grupo = BuscaArray("mnuPDT0601")
    frmManPDT0601.Show
    pSetFocus frmManPDT0601
End Sub

Private Sub mnuPerfiles_Click()
    frmManPerfiles.Grupo = BuscaArray("mnuPerfiles")
    frmManPerfiles.Show
    pSetFocus frmManPerfiles
End Sub

Private Sub mnuPeriodo_Click()
    frmManAnio.Grupo = BuscaArray("mnuPeriodo")
    frmManAnio.Show
    frmManAnio.gsAnioForm = gsAnio
    frmManAnio.tdbnAnio = gsAnio
    frmManAnio.tdbcEmpresa.BoundText = gsEmpresa
    pSetFocus frmManAnio
End Sub

Private Sub mnuPlanCuentas_Click()
    frmManPlanCuentas.Grupo = BuscaArray("mnuPlanCuentas")
    frmManPlanCuentas.Show
    pSetFocus frmManPlanCuentas
    
End Sub

Private Sub mnuPlantillaBalanceGeneral_Click()
    frmManPlantillaBalance.Grupo = BuscaArray("mnuPlantillaBalanceGeneral")
    frmManPlantillaBalance.Show
    pSetFocus frmManPlantillaBalance
    
End Sub

Private Sub mnuPLE_Click()
    FrmExportPLe.Show
End Sub

Private Sub mnuRatiosFinancieros_Click()
    frmPrcRatios.Grupo = BuscaArray("mnuRatiosFinancieros")
    frmPrcRatios.Show
    pSetFocus frmPrcRatios
End Sub


Private Sub MnuRegistroAsientos_Click()
    On Error GoTo serror
    
    If Len(Trim(CuentaCfgAuto("SEL_GAN"))) = 0 Then
       Mensajes "No se tiene definido la Cuenta Ganancia por Diferencia de Cambio", vbInformation
       Exit Sub
    End If
    
    If Len(Trim(CuentaCfgAuto("SEL_PER"))) = 0 Then
       Mensajes "No se tiene definido la Cuenta Perdida por Diferencia de Cambio", vbInformation
       Exit Sub
    End If

    With frmManAsientosContables
        .Grupo = BuscaArray("MnuRegistroAsientos")
        '.Show
        Load (frmManAsientosContables)
        .ZOrder 0
    End With
    Exit Sub
    
serror:
End Sub

Private Sub mnuRegistroEstractoBancario_Click()
    frmManEstractoBancario.Grupo = BuscaArray("mnuRegistroEstractoBancario")
    frmManEstractoBancario.Show
    pSetFocus frmManEstractoBancario
End Sub

Private Sub mnuRegistroPresupuesto_Click()
    frmManPresupuestos.Grupo = BuscaArray("mnuRegistroPresupuesto")
    frmManPresupuestos.Show
    pSetFocus frmManPresupuestos
End Sub

Private Sub MnuRepCostCta_Click()
frmRepCostosPorCuenta.Show
End Sub

Private Sub MnuRepCostoCta_Click()

End Sub

Private Sub mnuRepDetracciones_Click()
    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
    
    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuRepDetracciones")
    frmRepAnexoInvBalance.TituloSunat = "Reporte de Detracciones"
    frmRepAnexoInvBalance.Show
    frmRepAnexoInvBalance.ReporteSunat = "Detrac"
    frmRepAnexoInvBalance.CuentaInvBal = ""
    frmRepAnexoInvBalance.fraHastaMes.Visible = True
    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuReporteAnalisisLibro_Click()
    frmRepAsientosxLibros.Grupo = BuscaArray("mnuReporteAnalisisLibro")
    frmRepAsientosxLibros.Show
    pSetFocus frmRepAsientosxLibros
End Sub

Private Sub mnuReporteAsientosEntidades_Click()
    frmRepAnaliticoProveedores.Grupo = BuscaArray("mnuReporteAsientosEntidades")
    frmRepAnaliticoProveedores.Show
    pSetFocus frmRepAnaliticoProveedores
End Sub

Private Sub mnuReporteDaot_Click()
    frmRepDaot.Grupo = BuscaArray("mnuReporteDaot")
    frmRepDaot.Show
    pSetFocus frmRepDaot
End Sub

Private Sub mnuReporteEjecucionPresupuesto_Click()
    frmRepPresupuestoEjecucion.Grupo = BuscaArray("mnuReporteEjecucionPresupuesto")
    frmRepPresupuestoEjecucion.Show
    pSetFocus frmRepPresupuestoEjecucion
End Sub


Private Sub mnuReporteRegAuxiliares_Click()
    FrmRepAuxiliarVentas.gsGrupo = BuscaArray("mnu_reg_aux_ventas")
    FrmRepAuxiliarVentas.Show
    pSetFocus FrmRepAuxiliarVentas
End Sub

Private Sub mnuRestaurarBaseDatos_Click()
    frmPrcRestore.Grupo = BuscaArray("mnuRestaurarBaseDatos")
    frmPrcRestore.Show
    pSetFocus frmPrcRestore
End Sub

Private Sub mnuResumenCentroCosto_Click()
    frmRepResumenCentrocosto.Grupo = BuscaArray("mnuResumenCentroCosto")
    frmRepResumenCentrocosto.Show
    pSetFocus frmRepResumenCentrocosto
End Sub

Private Sub mnuResumenCentroCostoMes_Click()
    frmRepResumenCentroCostoMes.Grupo = BuscaArray("mnuResumenCentroCostoMes")
    frmRepResumenCentroCostoMes.Show
    pSetFocus frmRepResumenCentroCostoMes
End Sub

Private Sub mnuSaldosCuenta_Click()
    frmRepSaldosCuenta.Show
    frmRepSaldosCuenta.Grupo = BuscaArray("mnuSaldosCuenta")
    pSetFocus frmRepSaldosCuenta
End Sub

Private Sub mnuSaldosNetos_Click()
    frmRepSaldosNetos.Grupo = BuscaArray("mnuSaldosNetos")
    frmRepSaldosNetos.Show
    pSetFocus frmRepSaldosNetos
End Sub

Private Sub mnuSalir_Click()
    Salida = True
    Unload Me
End Sub

Private Sub mnuSeguimientoCheques_Click()
    frmRepSeguimientoCheques.Grupo = BuscaArray("mnuSeguimientoCheques")
    frmRepSeguimientoCheques.Show
    pSetFocus frmRepSeguimientoCheques
    
End Sub

Private Sub mnuTablasMnt_Click()
    frmManTablas.Grupo = BuscaArray("mnuTablasMnt")
    frmManTablas.Show
    pSetFocus frmManTablas
End Sub

Private Sub mnuTipoCambio_Click()
    frmManTipoCambio.Grupo = BuscaArray("mnuTipoCambio")
    frmManTipoCambio.Show
    pSetFocus frmManTipoCambio
End Sub

Private Sub mnuTipoDocumento_Click()
    frmManTipoDocumento.Grupo = BuscaArray("mnuTipoDocumento")
    frmManTipoDocumento.Show
    pSetFocus frmManTipoDocumento
End Sub

Private Sub mnuTipoEntidad_Click()
    frmManTipoEntidad.Grupo = BuscaArray("mnuTipoEntidad")
    frmManTipoEntidad.Show
    pSetFocus frmManTipoEntidad
End Sub

Private Sub mnuTipoLibro_Click()
    frmManLibros.Grupo = BuscaArray("mnuTipoLibro")
    frmManLibros.Show
    pSetFocus frmManLibros
End Sub

Private Sub mnuTipoMoneda_Click()
    frmManTipoMoneda.Grupo = BuscaArray("mnuTipoMoneda")
    frmManTipoMoneda.Show
    pSetFocus frmManTipoMoneda
End Sub

Private Sub mnuTiposAsiento_Click()
    frmManPlantillaTipoAsiento.Grupo = BuscaArray("mnuTiposAsiento")
    frmManPlantillaTipoAsiento.Show
    pSetFocus frmManPlantillaTipoAsiento
End Sub

Private Sub mnuUpdLibroElec_Click()
    frmPrcActualizaEstadosPLE.Show
End Sub

Private Sub mnuUsuarios_Click()
    frmManPerfilUsuario.Grupo = BuscaArray("mnuUsuarios")
    frmManPerfilUsuario.Show
    pSetFocus frmManPerfilUsuario
End Sub

Private Sub mnuValores_Click()
'    FrmManRegAcc.Show
'    Exit Sub
    FrmManRegAcc.Grupo = BuscaArray("FrmManRegAcc")
    FrmManRegAcc.Show
    pSetFocus FrmManRegAcc
End Sub

Private Sub mnuVerAsientosImportados_Click()
    frmPrcEliminaImportaciones.Grupo = BuscaArray("mnuVerAsientosImportados")
    frmPrcEliminaImportaciones.Show
    pSetFocus frmPrcEliminaImportaciones
End Sub

Private Sub regPercion_Click()
    frmRepRegistroRetencion.Grupo = BuscaArray("regPercion")
    frmRepRegistroRetencion.Show
    pSetFocus frmRepRegistroRetencion
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub m_mes_Click(Index As Integer)
    gsPeriodo = Right("00" & CStr(Index), 2)
    Me.stbMdi.Panels(2).Text = "  " & NombreMes(gsPeriodo) & " DEL " & gsAnio & "  "
    
    GrabaPeriodoActivo
End Sub

Private Sub CheckMenuMes()
    Dim i As Integer
    For i = 0 To 14
        m_mes(i).Checked = False
    Next i
    m_mes(gsPeriodo).Checked = True
End Sub

Private Sub stbMdi_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (gsInicioVoucher = False) And ActivaMenuEmergente = True Then
       CheckMenuMes
       gsMesSistema = m_Meses
       PopupMenu m_Meses, vbPopupMenuLeftAlign
    End If
    
End Sub

Private Sub stbMdi_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 2 Then
        ActivaMenuEmergente = True
    Else
        ActivaMenuEmergente = False
    End If
End Sub

Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Screen.MousePointer = vbHourglass
    Select Case Button.Index
        Case 1:
            If mnuPlanCuentas.Visible = True And mnuTablas.Visible = True Then
                tbrOpciones.Buttons(Button.Index).Enabled = False
                DoEvents
                mnuPlanCuentas_Click
                DoEvents
                tbrOpciones.Buttons(Button.Index).Enabled = True
            Else
                Mensajes "Esta opción esta restringida para el usuario", vbExclamation + vbOKOnly
            End If
            Screen.MousePointer = vbNormal
        Case 2:
            If mnuTipoCambio.Visible = True And mnuTablas.Visible = True Then
                tbrOpciones.Buttons(Button.Index).Enabled = False
                DoEvents
                mnuTipoCambio_Click
                DoEvents
                tbrOpciones.Buttons(Button.Index).Enabled = True
                
            Else
                Mensajes "Esta opción esta restringida para el usuario", vbExclamation + vbOKOnly
            End If
            Screen.MousePointer = vbNormal
        Case 3:
            If mnuEntidadesMnt.Visible = True And mnuTablas.Visible = True Then
                tbrOpciones.Buttons(Button.Index).Enabled = False
                DoEvents
                mnuEntidadesMnt_Click
                DoEvents
                tbrOpciones.Buttons(Button.Index).Enabled = True
                
            Else
                Mensajes "Esta opción esta restringida para el usuario", vbExclamation + vbOKOnly
            End If
            Screen.MousePointer = vbNormal
        Case 4:
        
            'SEPARADOR
        Case 5:
            If MnuRegistroAsientos.Visible = True And mnuIngresos.Visible = True Then
                tbrOpciones.Buttons(Button.Index).Enabled = False
                DoEvents
                MnuRegistroAsientos_Click
                DoEvents
                Me.tbrOpciones.Buttons(Button.Index).Enabled = True
                
            Else
                Mensajes "Esta opción esta restringida para el usuario", vbExclamation + vbOKOnly
            End If
            Screen.MousePointer = vbNormal
        Case 6:
            If mnuConfigOperac.Visible = True And mnuTablas.Visible = True Then
                tbrOpciones.Buttons(Button.Index).Enabled = False
                DoEvents
            
                mnuConfigOperac_Click
                DoEvents
                tbrOpciones.Buttons(Button.Index).Enabled = True
                
            Else
                Mensajes "Esta opción esta restringida para el usuario", vbExclamation + vbOKOnly
            End If
            Screen.MousePointer = vbNormal
        Case 7:
            'SEPARADOR
        Case 8:
            mnuWindowCascade_Click
            Screen.MousePointer = vbNormal
        Case 9:
            mnuWindowTileHorizontal_Click
            Screen.MousePointer = vbNormal
        Case 10:
            mnuWindowArrangeIcons_Click
            Screen.MousePointer = vbNormal
        Case 11:
            'SEPARADOR
        Case 12:
            Screen.MousePointer = vbNormal
            mnuCambioAnio_Click
        Case 13:
            Screen.MousePointer = vbNormal
            Logoff = True
            mnuSalir_Click
        Case 14:
            'SEPARADOR
        Case 15:
            Call LlamarAyuda
            Screen.MousePointer = vbNormal
        Case 17:
            Screen.MousePointer = vbNormal
            mnuSalir_Click
            Exit Sub
    End Select
    
End Sub

Private Sub LlamarAyuda()
    'If mnuAyuda.Visible = True Then
        If Not Me.ActiveForm Is Nothing Then
           Ayuda Me.ActiveForm.Name
        Else
           ShowContents 1
        End If
    'End If

End Sub

Private Function GrupoUsuario(usuario As String) As String
    ' *** Validar el ingreso al sistema
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q cuenta exista
    Set clDatos = New clsMantoTablas
    sqlSp = "spSg_GrabaUsuarios 'SEL_REG_EMP', '" & usuario & "','', '', '', '', '', '" & gsEmpresa & "', '' ,'" & gsSOFT & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        GrupoUsuario = ""
    Else
        GrupoUsuario = CE(rsArreglo("usu_cGrupo"))
    End If
    CerrarRecordSet rsArreglo
    Set clDatos = Nothing
    ' ***
End Function

Private Sub mnuAcerca_Click()
    frmAcercaDe.Show
    frmAcercaDe.ZOrder 0
End Sub

Private Sub mnuContenido_Click()
    ShowContents 1
End Sub

Private Sub mnuIndice_Click()
    ShowIndex 1
End Sub

Private Sub mnuBusqueda_Click()
    ShowSearch 1
    ShowTopicID 1, 2
End Sub

Public Function BuscaForm(Nombre As String) As Boolean
    On Error GoTo serror
    Dim i As Integer
    BuscaForm = False

    For i = 1 To Forms.Count - 1
        If Forms(i).Name = Nombre Then
           BuscaForm = True
           Exit For
        End If
    Next i
    Exit Function
serror:
End Function

Public Function BuscaFormTag(cNombre As String, cTag As String) As Boolean
    On Error GoTo serror
    Dim i As Integer
    BuscaFormTag = False
    
    For i = 1 To Forms.Count - 1
        If Forms(i).Name = cNombre And Forms(i).Tag = cTag Then
        
           'Unload Forms(i)
           Forms(i).ZOrder 0
           BuscaFormTag = True
           Exit For
        End If
    Next i
    Exit Function
serror:
End Function

Private Sub mnuF0309_33_Click()
    Call GenerarVentanaAntigua("F0309_33", "Formato : " & mnuF0309_33.Caption, "33")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_33.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_33"
'    frmRepAnexoInvBalance.CuentaInvBal = "33"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0309_34_Click()
    Call GenerarVentanaAntigua("F0309_34", "Formato : " & mnuF0309_34.Caption, "34")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_34.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_34"
'    frmRepAnexoInvBalance.CuentaInvBal = "34"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0309_35_Click()
    Call GenerarVentanaAntigua("F0309_35", "Formato : " & mnuF0309_35.Caption, "35")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_35.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_35"
'    frmRepAnexoInvBalance.CuentaInvBal = "35"
'    pSetFocus frmRepAnexoInvBalance

End Sub

Private Sub mnuF0309_38_Click()
    Call GenerarVentanaAntigua("F0309_38", "Formato : " & mnuF0309_38.Caption, "38")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_38.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_38"
'    frmRepAnexoInvBalance.CuentaInvBal = "38"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0309_39_Click()
    Call GenerarVentanaAntigua("F0309_39", "Formato : " & mnuF0309_39.Caption, "39")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_39.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_39"
'    frmRepAnexoInvBalance.CuentaInvBal = "39"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub mnuF0309_45_Click()
    Call GenerarVentanaAntigua("F0309_45", "Formato : " & mnuF0309_45.Caption, "45")
    
'    If BuscaForm("frmRepAnexoInvBalance") Then frmRepAnexoInvBalance.CerrarForm
'
'    frmRepAnexoInvBalance.Grupo = BuscaArray("mnuAnexoInvBalance")
'    frmRepAnexoInvBalance.TituloSunat = "Formato : " & mnuF0309_45.Caption
'    frmRepAnexoInvBalance.Show
'    frmRepAnexoInvBalance.ReporteSunat = "F0309_45"
'    frmRepAnexoInvBalance.CuentaInvBal = "45"
'    pSetFocus frmRepAnexoInvBalance
End Sub

Private Sub GenerarVentanaAntiguaBancos(cReporteEmp As String, cCadena As String, cCtaInvBal As String)
    If BuscaFormTag("frmRepLibroBancos", cReporteEmp) = False Then
        Set oReportesEmpAntBcos = New frmRepLibroBancos
        With oReportesEmpAntBcos
            .ReporteSunat = cReporteEmp
            .TituloSunat = cCadena
            .Caption = cCadena
            .Tag = cReporteEmp
            .Grupo = BuscaArray("mnu" & cReporteEmp)
            .Show
            
            .ZOrder 0
        End With
    End If
End Sub

Private Sub GenerarVentanaAntigua(cReporteEmp As String, cCadena As String, cCtaInvBal As String)
    If BuscaFormTag("frmRepAnexoInvBalance", cReporteEmp) = False Then
        Set oReportesEmpAnt = New frmRepAnexoInvBalance
        With oReportesEmpAnt

            .ReporteSunat = cReporteEmp
            .TituloSunat = cCadena
            .Caption = cCadena
            .Tag = cReporteEmp
            .Grupo = BuscaArray("mnu" & cReporteEmp)
            
            If cCtaInvBal <> "" Then
                On Error Resume Next
                .Label3(0).Caption = "MONEDA"
                .Label3(1).Caption = "MONEDA"
            End If
            
            If InStr(1, cCadena, "General") <> 0 Or InStr(1, cCadena, "Ganancia") <> 0 Then
                '.ChkVerMes.Visible = True
            End If
            
            .Show
            .CuentaInvBal = cCtaInvBal
            .ZOrder 0
        End With
    End If
End Sub

Private Sub GenerarVentana(cReporteEmp As String, cCadena As String)
    If BuscaFormTag("frmRepAnexoInvBalanceEmp", cReporteEmp) = False Then
        Set oReportesEmp = New frmRepAnexoInvBalanceEmp
        With oReportesEmp
            .ReporteSunat = cReporteEmp
            .Caption = cCadena
            .TituloSunat = cCadena
            .Tag = cReporteEmp
            .Grupo = BuscaArray("mnu" & cReporteEmp)
            .Show
            If cReporteEmp = "PCGE50_0" Or cReporteEmp = "PCGE50_1" Then
                cReporteEmp = "50"
            End If
            .CuentaInvBal = Right(cReporteEmp, 2)
            .ZOrder 0
        End With
    End If
End Sub

Private Sub mnuPCGE10_Click()
    Call GenerarVentana("PCGE10", mnuPCGE10.Caption)
End Sub

Private Sub mnuPCGE11_Click()
    Call GenerarVentana("PCGE11", mnuPCGE11.Caption)
End Sub

Private Sub mnuPCGE12_Click()
    Call GenerarVentana("PCGE12", mnuPCGE12.Caption)
End Sub

Private Sub mnuPCGE13_Click()
    Call GenerarVentana("PCGE13", mnuPCGE13.Caption)
End Sub

Private Sub mnuPCGE14_Click()
    Call GenerarVentana("PCGE14", mnuPCGE14.Caption)
End Sub

Private Sub mnuPCGE16_Click()
    Call GenerarVentana("PCGE16", mnuPCGE16.Caption)
End Sub

Private Sub mnuPCGE17_Click()
    Call GenerarVentana("PCGE17", mnuPCGE17.Caption)
End Sub

Private Sub mnuPCGE18_Click()
    Call GenerarVentana("PCGE18", mnuPCGE18.Caption)
End Sub

Private Sub mnuPCGE19_Click()
    Call GenerarVentana("PCGE19", mnuPCGE19.Caption)
End Sub

Private Sub mnuPCGE20_Click()
    Call GenerarVentana("PCGE20", mnuPCGE20.Caption)
End Sub

Private Sub mnuPCGE21_Click()
    Call GenerarVentana("PCGE21", mnuPCGE21.Caption)
End Sub

Private Sub mnuPCGE22_Click()
    Call GenerarVentana("PCGE22", mnuPCGE22.Caption)
End Sub

Private Sub mnuPCGE23_Click()
    Call GenerarVentana("PCGE23", mnuPCGE23.Caption)
End Sub

Private Sub mnuPCGE24_Click()
    Call GenerarVentana("PCGE24", mnuPCGE24.Caption)
End Sub

Private Sub mnuPCGE25_Click()
    Call GenerarVentana("PCGE25", mnuPCGE25.Caption)
End Sub

Private Sub mnuPCGE26_Click()
    Call GenerarVentana("PCGE26", mnuPCGE26.Caption)
End Sub

Private Sub mnuPCGE27_Click()
    Call GenerarVentana("PCGE27", mnuPCGE27.Caption)
End Sub

Private Sub mnuPCGE28_Click()
    Call GenerarVentana("PCGE28", mnuPCGE28.Caption)
End Sub

Private Sub mnuPCGE29_Click()
    Call GenerarVentana("PCGE29", mnuPCGE29.Caption)
End Sub

Private Sub mnuPCGE30_Click()
    Call GenerarVentana("PCGE30", mnuPCGE30.Caption)
End Sub

Private Sub mnuPCGE31_Click()
    Call GenerarVentana("PCGE31", mnuPCGE31.Caption)
End Sub

Private Sub mnuPCGE32_Click()
    Call GenerarVentana("PCGE32", mnuPCGE32.Caption)
End Sub

Private Sub mnuPCGE33_Click()
    Call GenerarVentana("PCGE33", mnuPCGE33.Caption)
End Sub

Private Sub mnuPCGE34_Click()
    Call GenerarVentana("PCGE34", mnuPCGE34.Caption)
End Sub

Private Sub mnuPCGE35_Click()
    Call GenerarVentana("PCGE35", mnuPCGE35.Caption)
End Sub

Private Sub mnuPCGE36_Click()
    Call GenerarVentana("PCGE36", mnuPCGE36.Caption)
End Sub

Private Sub mnuPCGE37_Click()
    Call GenerarVentana("PCGE37", mnuPCGE37.Caption)
End Sub

Private Sub mnuPCGE38_Click()
    Call GenerarVentana("PCGE38", mnuPCGE38.Caption)
End Sub

Private Sub mnuPCGE39_Click()
    Call GenerarVentana("PCGE39", mnuPCGE39.Caption)
End Sub

Private Sub VentanaAbrir(ByRef oForm As Form)
    Dim i As Integer, x As Integer
    '*** BUSCAR SIS EL FORMULARIO YA ESTA ABIERTO
    For i = 0 To Forms.Count - 1
        If Forms(i).hwnd = oForm.hwnd Then
            '*** MOSTRAR SOLO SI ESTA HABILITADO
            If Forms(i).Enabled = True Then
                oForm.WindowState = vbNormal
                oForm.ZOrder 0
                '*** ACTIVAR EL TAB
                tsForms.DeselectAll
                For x = 1 To tsForms.Tabs.Count
                    If tsForms.Tabs(x).key = "frm" & CE(oForm.hwnd) Then
                        tsForms.Tabs(x).Selected = True
                        Exit For
                    End If
                Next
            End If
            Exit Sub
        End If
    Next
    '*** ABRIR Y MOSTRAR FORMULARIO SI NO EXISTE
    oForm.Show
    oForm.WindowState = vbNormal
    oForm.ZOrder 0
End Sub

'*** REDIMENSIONAR EL TAB
Private Sub TabForms_Redim()
On Error GoTo errHand
    
    'tsForms.Left = Me.Width + 30
    tsForms.Width = Me.Width - 200 '- tsForms.Left - picBarraDerecha.Width
Exit Sub
errHand:
End Sub

Private Sub tsForms_Click()
Dim i As Integer
On Error GoTo errHand
    If tsForms.Tabs.Count < 1 Then Exit Sub
    For i = 0 To Forms.Count - 1
        If "frm" & Forms(i).hwnd = tsForms.SelectedItem.key Then
            Forms(i).ZOrder
            Exit For
        End If
    Next
Exit Sub
errHand:

End Sub

Public Sub TabForm_CrearN(ByVal nFormHwnd As Long, Optional ByVal nMainHwnd As Long = 0)
Dim i As Integer, x As Integer
Dim bExiste As Boolean
On Error GoTo errHand
    '*** UBICAR EL FORMULARIO EN FORMS
    For x = 0 To Forms.Count - 1
        If Forms(x).hwnd = nFormHwnd Then bExiste = True: Exit For
    Next
    If bExiste Then
        tsForms.DeselectAll
        For i = 1 To tsForms.Tabs.Count
            '*** SI ES FORMULARIO DEPENDIENTE
            If nMainHwnd <> 0 Then
                If tsForms.Tabs(i).key = "frm" & CE(nMainHwnd) Then
                    tsForms.Tabs(i).Selected = True
                    tsForms.Tabs(i).Caption = CE(Forms(x).Caption)
                    tsForms.Tabs(i).key = "frm" & CE(Forms(x).hwnd)
                    Exit Sub
                End If
            Else
                If tsForms.Tabs(i).key = "frm" & CE(Forms(x).hwnd) Then
                    tsForms.Tabs(i).Selected = True
                    Exit Sub
                End If
            End If
        Next
        tsForms.Tabs.Add , "frm" & CE(Forms(x).hwnd), Forms(x).Caption, IIf(Forms(x).Name = "frmReportPreview", 2, 1)
        tsForms.Tabs(tsForms.Tabs.Count).Selected = True
    End If
Exit Sub
errHand:
End Sub

'*** CERRAR EL TAB DEL FORMULARIO AL CERRARLO
Public Sub TabForm_CerrarN(ByVal nFormHwnd As Long, Optional ByVal nMainHwnd As Long = 0)
Dim i As Integer, x As Integer
Dim bExiste As Boolean
On Error GoTo errHand
    If nMainHwnd <> 0 Then
        '*** UBICAR EL MAIN EN FORMS
        For x = 0 To Forms.Count - 1
            If Forms(x).hwnd = nMainHwnd Then Forms(x).Enabled = True: bExiste = True: Exit For
        Next
    End If
    For i = 1 To tsForms.Tabs.Count
        If bExiste Then
            If tsForms.Tabs(i).key = "frm" & CE(nFormHwnd) Then
                tsForms.Tabs(i).Selected = True
                tsForms.Tabs(i).Caption = CE(Forms(x).Caption)
                tsForms.Tabs(i).key = "frm" & CE(nMainHwnd)
                Exit Sub
            End If
        Else
        '*** BUSCAR SI EXISTE Y BORRARLO
            If tsForms.Tabs(i).key = "frm" & CE(nFormHwnd) Then
                tsForms.Tabs.Remove i
                Exit Sub
            End If
        End If
    Next
Exit Sub
errHand:
End Sub

Private Sub mnuConceptosLibros_Click()
    'Mensajes "Esta opcion se habilitará en la nueva version"
    frmManConceptoLibros.Grupo = BuscaArray("mnuConceptosLibros")
    frmManConceptoLibros.Show
    pSetFocus frmManConceptoLibros

End Sub

Private Sub tsForms_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then 'botom derecho del mouse
        mnuCerrarTabVentana.Caption = "Cerrar : " & BuscarCaptionFormulario(tsForms.SelectedItem.key)
        PopupMenu mnuCerrarTab, vbPopupMenuLeftAlign
    End If
End Sub

Private Function BuscarCaptionFormulario(cKey As String) As String
Dim i As Integer
BuscarCaptionFormulario = ""
On Error GoTo errHand
    If tsForms.Tabs.Count < 1 Then Exit Function
    For i = 0 To Forms.Count - 1
        If "frm" & Forms(i).hwnd = cKey Then
            BuscarCaptionFormulario = Forms(i).Caption
            Exit For
        End If
    Next
Exit Function
errHand:
BuscarCaptionFormulario = ""
End Function

Private Sub mnuCerrarTabVentana_Click()
Dim i As Integer
On Error GoTo errHand
    If tsForms.Tabs.Count < 1 Then Exit Sub
    For i = 0 To Forms.Count - 1
        If "frm" & Forms(i).hwnd = tsForms.SelectedItem.key Then
            Unload Forms(i)
            Exit For
        End If
    Next
Exit Sub
errHand:
End Sub

Private Sub mnuGru_Agrupar_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Grupos_Agrupar
End Sub

Private Sub mnuGru_Desagrupar_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Grupos_Desagrupar
End Sub

Private Sub mnuGru_DesagTodo_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Grupos_DesagruparTodo
End Sub

Private Sub mnuGru_Copiar_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Grupos_Copiar
End Sub

Private Sub mnuGru_Pegar_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Grupos_Pegar
End Sub

Private Sub mnuAsigConcepto_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Concepto_Asignar
End Sub

Private Sub mnuQuitarConcepto_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Concepto_Quitar
End Sub
Private Sub mnuQuitarTodosConceptos_Click()
    If BuscaForm("frmManAsientosContables") Then frmManAsientosContables.Concepto_QuitarTodos
End Sub



