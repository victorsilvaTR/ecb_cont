Attribute VB_Name = "modDeclaraciones"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1

'--------------------------------
Global Const gsAMBOSPLANDECUENTAS = True
'--------------------------------
Global gsCambioEmpresa As Boolean
Global gsInstancia As String
Global nContadorProc  As Integer
Global gsPrnDriverName As String
Global gsPrnPrinterName As String
Global gsPrnPortName As String
Global gsNomTipoImp As Boolean
Global strNewPrinterName As String
Global sCuenta As String
Global sNroCol As Integer
Global gintBiMoneda As Integer
Global gintPercepcion As Integer
Global gintRetencion As Integer
Global gintTipoRegimen As Integer 'frt_rvie
Global gstrCuentaCostoVenta As String
Global gintLEVentaSimplificado As Integer
Global gintLECompraSimplificado As Integer
Global gstrVersionLE As String
Global gsRVIE As String 'frt_rvie
Global gstrCostoProduccion As String
Global lsFecha As Date
'--------------------------------
Global Const gsNombreModulo = "ECB-Cont"
'Global Const gsVersion = "   v2.6.5"
'Global Const gsVersion = "   v2.9.0"
''Global Const gsVersion = "   v2.13.6" 'frt_rvie
'Global Const gsVersion = "   v2.14.1" 'hlp20231010

Global Const gsVersion = "   v2.15.0" 'hlp20231010
Global Const gsNumSubida = 0          'pbp20250710
'-------------------------------------------------------------------------
Global Const gsSOFT = "001"
Global Const gsDSN = "dsnECB"
'-------------------------------------------------------------------------
Global gsImportacion As Boolean
'-------------------------------------------------------------------------
Global grstBusqueda As ADODB.Recordset  ' RECORDSET PARA BUSQUEDAS
'-------------------------------------------------------------------------
Global gsLetPagar As String
Global gsLetCobrar As String
Global gsCheque As String

'-------------------------------------------------------------------------
'-- campos de formulario de busqueda
Global gsCodigo As String
Global gsCadena As String
Global gsCampo3 As String
Global gsCampo4 As String
Global gsCampo5 As String
Global gsCampo6 As String
'-------------------------------------------------------------------------
Global gsDetalle As String
Global gsProvEnviada As Boolean
'-------------------------------------------------------------------------
Global gsDiarioSimplificado As Integer
Global gsKeyCodePress As Boolean
Global gsKeyPressF1 As Boolean
Global lsLibShow As Boolean
Global lsLibroCom As String
Global lsLibroVen As String
Global lsLibroHon As String
Global lsLibroApe As String
Global lsLibroDif As String
Global lsLibroRet As String
Global lsLibroTransferenciaCancelacion As String
Global lsLibroTransCancAutomatico As String
Global lsLibroAjusteNIF As String
Global lsLibroCierre As String
Global lsLibroCajIng As String
Global lsLibroCajEgr As String
Global lsLibroDiario As String
Global lsTipOperaLib As String
Global LibroDefault As String
Global gsBaseImpDefCom As String
Global gsBaseImpDefVtas As String
Global gsTipoPlan As Integer
Global gsPLE As Integer
Global gsNumDigDiarioSimpRep As Integer
Global liQuiebre   As Integer

Global gsTDNC As String
Global gcnSistema As New ADODB.Connection
Global gcnSistemaAdv As New ADODB.Connection
Global gsCadenaConexion As String

Global gsEmpresa As String
Global gsEmpresaNom As String
Global gsRutaBackup As String
Global gsGenDsn As String
Global gsUsuario As String
Global gsAdmin As String
Global gsRUC As String
Global gsBDUS As String
Global gsBDPW As String
Global gsBD As String
Global gsEncriptada  As String
Global gsGenLog As String
Global gsGenLogMov As String
Global gsGenLogConsulta As String
Global gsAnio As String
Global gsServidor As String
Global gsPeriodo As String
Global gsAutenticacion As String

Global gsByMoneda As Integer
Global gsError As Boolean
Global gsMonedaNac As String
Global gsMonedaExt As String
Global gsNombreMonedaNac As String
Global gsNombreMonedaExt As String
Global gsMonedaNacAbrev As String
Global gsMonedaExtAbrev As String
Global gsMesSistema As String

Global gsColumna As Integer
Global gsSucursal As String
Global gsInicio  As Boolean 'VARIABLE PARA FORMULARIO DE CAMBIO DE EMPRESA, SI EL SHOW ES AL INICIO O CUANDO YA SE INGRESO AL SISTEMA
Global gsInicioVoucher  As Boolean
Global gsInicioCtasxC As Boolean
Global gsInicioCtasxP As Boolean
Global gsKey As Integer
Global gsArray As XArrayDB
Global gsPeriodoCOA As String
Global gsCodigoEnt As String
Global gsCadenaErr As String

Global Const gsPrivilegioAdmin = "0000"
Global Const gsColorActivado = &HFFFFFF
Global Const gsColorDesactProv = &HC0E0FF
Global Const gsColorRegAux = &H91D9AF
Global Const gsColorCCTitulo = &H80C0FF
Global Const gsColorCCSTitulo = &HC0E0FF
Global Const gsColorDesactivado = &HEFD8C2

Global gsCuentaDifGan As String
Global gsCuentaDifPer As String
Global gsCuentaRedGan As String
Global gsCuentaRedPer As String
Global gsFilasAfectadas As Integer
Global NombreReporte As String
Global SwHoja As Boolean

Global EstadoOri As String 'Indica el estado origen para el registro del asiento -> libro eletronico
Global EstadoDes As String 'Indica el estado posterior al modificar el asiento -> libro eletronico
Global EstadoDesTMP As String 'Indica el estado posterior al modificar el asiento, usado para el estado anulado guarda el estado ant. -> libro eletronico
Global SW_ActPLE As String 'Indica si sistema cuenta con PLE=1 o sin PLE = 0
Global EstadoLDOri As String 'Indica el estado origen para el registro de la cuenta-> libro Diario Detalle cuenta
Global EstadoLDDes As String 'Indica el estado posterior al modificar la cuenta -> libro Diario Detalle cuenta



Global tipoDocRef As String
Global SerieDocRef As String
Global NumDocRef As String
Global fechDocRef As String
Global ValEmb As Double
Global ValEmbAnt As Double

Global DocDepo As String
Global fechaDepo As String

'--------------------------------

Public Enum Tipo_Pagina
    defecto = 0
    CARTA = 1
    USA = 2
    OFICIO = 3
    A4 = 4
End Enum

Public Enum Orientacion_Pagina
    defecto = 0
    Vertical = 1
    Horizontal = 2
End Enum

Public Enum Tipo_Cambio
     TCM_PROMEDIO = 0
     TCM_COMPRA = 1
     TCM_VENTA = 2
End Enum

Public Enum Tipo
     CONSULTA = 0
     EDICION = 1
End Enum

Public Enum TipoGrupo
     G_CONSULTAR = 0
     G_INGRESAR = 1
     G_MODIFICAR = 2
     G_ELIMINAR = 3
End Enum

Public Enum TipoControl
     NINGUNO = 0
     PROCESOCOSTOS = 1
     BASEIMPONIBLE = 2
End Enum

Public Enum TipoDato
     X_CUENTA = 1
     X_CODIGO = 2
     X_ENTIDAD = 3
     X_TIPO = 4
     X_SERIE = 5
     X_NUMERO = 6
End Enum

Public Enum TipoImporte
     X_NACD = 1
     X_NACH = 2
     X_EXTD = 3
     X_EXTH = 4
     X_TC = 5
     X_FLAM = 6
     X_CORR = 7
End Enum

Public Enum TipoOtros
    X_CC = 1
    X_FD = 2
    X_FV = 3
    X_MO = 4
    X_OP = 5
    X_TM = 6
    X_MC = 7
    X_PC = 8
    X_DH = 9
    X_FM = 10
End Enum

'HT : 20091111
Global gsAccionRep   As Integer
Global giBold        As Integer
Global Gi_FlagImpresion As Integer
Global giCopias      As Integer
Global giLineas      As Integer
Global giEspacios    As Integer
Global Gs_HoraServ As String
Global gsPagina      As Long
Global gsPaginaPrincipal As Integer
Global gsConTotalPaginas As Long
Global gsControlPag As Boolean
Global gsCodMoneda As String
Global gsDiarioPeriodo As String
Global gsLinea       As String
Global gsTipoImp As String

'HT:20091226
Global gsMesRep As String
Global gsCtaIni As String
Global gsCtaFin As String
Global gsNombreVista As String

Global gsLdMesIni As String
Global gsLdMesFin As String
Global gsLdFechIni As String
Global gsLdFechFin As String

Public Gs_TamPapel As String
Public GsDestino As String

Public Gs_DesdePag   As Integer
Public Gs_HastaPag   As Integer

Public xGs_DesdePag   As Integer
Public xGs_HastaPag   As Integer
Public xGs_Principal  As Integer

'HT : 20100115
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global gsVersionWindowsMayor As Integer
Global gsVersionWindowsMenor As Integer

Public sSql As String
