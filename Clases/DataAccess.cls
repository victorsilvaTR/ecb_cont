VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
' DECLARACIONES GENERALES
'**************************************************************************

' *** Timers para la Apertura de la Conexion
Public Enum TEL_CONNECT_TIMERS_ENUM
    TEL_CONNECT_SHORT_TIMEOUT = 50000
    TEL_CONNECT_NORMAL_TIMEOUT = 60000
    TEL_CONNECT_LONG_TIMEOUT = 2000000
End Enum

' *** Tipos de Bloqueo para la Apertura de Recordsets
Public Enum TEL_LOCK_TYPES
    TEL_LOCK_READ_ONLY = ADODB.adLockReadOnly
    TEL_LOCK_OPTIMISTIC = ADODB.adLockOptimistic
    TEL_LOCK_PESSIMISTIC = ADODB.adLockPessimistic
    TEL_LOCK_BATCH_OPTIMISTIC = ADODB.adLockBatchOptimistic
End Enum

' *** Ubicacion del Cursor
Public Enum TEL_CURSOR_LOCATION
    TEL_CURSOR_NONE = ADODB.adUseNone
    TEL_CURSOR_CLIENT = ADODB.adUseClient
    TEL_CURSOR_SERVER = ADODB.adUseServer
End Enum

' *** Tipos de Cursores para Recordsets
Public Enum TEL_CURSOR_TYPE
    TEL_TYPE_FORWARD_ONLY = ADODB.adOpenForwardOnly
    TEL_TYPE_KEYSET = ADODB.adOpenKeyset
    TEL_TYPE_DYNAMIC = ADODB.adOpenDynamic
    TEL_TYPE_STATIC = ADODB.adOpenStatic
End Enum

' *** Columnas del Buffer de Parametros
Public Enum TEL_SP_PARAM_BUFFER
    TEL_SP_PARAMET_NAME = 0
    TEL_SP_PARAMET_TYPE = 1
    TEL_SP_PARAMET_DIRECTION = 2
    TEL_SP_PARAMET_SIZE = 3
    TEL_SP_PARAMET_VALUE = 4
End Enum

' *** Direcciones de los Parametros
Public Enum TEL_SP_PARAM_DIRECTION
    TEL_PARAM_INPUT = ADODB.adParamInput
    TEL_PARAM_OUTPUT = ADODB.adParamOutput
    TEL_PARAM_INPUT_OUTPUT = ADODB.adParamInputOutput
    TEL_PARAM_RETURN_VALUE = ADODB.adParamReturnValue
End Enum

Private Const msOBJECT_NAME As String = "IDataAccess"
' *** Variables para Albergar Objetos de ADO
Private moCachedConnection As ADODB.Connection

Public Property Get Id() As String
    Id = msOBJECT_NAME
End Property

Private Sub Class_Terminate()
    ' *** Cierra Cualquier Conexion Pendiente
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
        Set moCachedConnection = Nothing
    End If
End Sub

'**************************************************************************
' Nombre....: CacheConnection()
' Proposito.: Esta funcion abre una conexion de ADO y la mantiene Abierta.
' Inputs....: Opcional Cadena de Connección y Valor para ConnectTimeOut.
' Retorno...: Retorna True si la Función es Satisfactoria
' Efectos...: Cuando es llamada sin argumentos CIERRA la Conexion.
'**************************************************************************

Public Function CacheConnection(Optional ByVal sConnect As String) As Boolean

    On Error GoTo ErrorHandler
    Dim lErrNo      As Long
    Dim sErrDesc    As String
    Dim sErrSrc     As String
    
    'Asumir Falla en la Conexion

    CacheConnection = False
    
    'Cerrar Cualquier Conexion Existente
    
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
        Set moCachedConnection = Nothing
    End If

    'Si el Parametro no es Nulo, Entonces realizar la Conexión
    
    If sConnect <> vbNullString Then
        Set moCachedConnection = New ADODB.Connection
        'moCachedConnection.ConnectionTimeout = TEL_CONNECT_NORMAL_TIMEOUT
        moCachedConnection.Open sConnect
    End If
    
    CacheConnection = True
    
    Exit Function
    
ErrorHandler:
                        'Valores de Retorno de Error
            lErrNo = Err.Number
            sErrDesc = Err.Description
            sErrSrc = Err.Source
            Set moCachedConnection = Nothing
End Function

'**************************************************************************
' Nombre....: DestroyConecction()
' Proposito.: Esta Función Destruye la Conexión
' Inputs....: Ninguno
' Se asume..: Ninguno
' Retorno...: Retorna True si la Función es Satisfactoria
' Efectos...: Ninguno
'**************************************************************************
Public Function DestroyConecction() As Boolean

    On Error GoTo ErrorHandler
    Dim lErrNo      As Long
    Dim sErrDesc    As String
    Dim sErrSrc     As String
    
    ' Asumir Falla en la Función
    
    DestroyConecction = False
    
    'Cerrar Cualquier Conexion Existente
    
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.Close
    End If

    Set moCachedConnection = Nothing
    
    DestroyConecction = True
    
    Exit Function
    
ErrorHandler:
            'Valores de Retorno de Error
            lErrNo = Err.Number
            sErrDesc = Err.Description
            sErrSrc = Err.Source
            
            Set moCachedConnection = Nothing
End Function

'**************************************************************************
' Nombre....: IniciarTransaccion()
' Proposito.: Realiza un Begin en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function IniciarTransaccion()
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.BeginTrans
    End If
End Function

'**************************************************************************
' Nombre....: FinalizarTransaccion()
' Proposito.: Realiza un Commit en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function FinalizarTransaccion()
On Error GoTo ERROR:
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.CommitTrans
    End If
ERROR:
End Function

'**************************************************************************
' Nombre....: CancelarTransaccion()
' Proposito.: Realiza un RollBack en la Transacción
' Inputs....: N/A
' Se asume..: N/A
' Retorno...: N/A
' Efectos...: N/A
'**************************************************************************
Public Function CancelarTransaccion()
    On Error GoTo ERROR:
    If Not moCachedConnection Is Nothing Then
        moCachedConnection.RollbackTrans
    End If
    Exit Function
ERROR:
    
End Function

'**************************************************************************
' Nombre....: ExecuteActionSP()
' Proposito.: Ejecuta un sp que no devuelve filas (Query de acciones)
' Inputs....: Stored Procedure, Array de Parametros
' Se asume..: N/A
' Retorno...: Retorna True si la función tiene exito
' Efectos...: Retorna en el mismo arreglo de parametros los valores
'             de los parametros del SP definidos como OUTPUT.
'**************************************************************************
Public Function ExecuteActionSP(ByVal sStoredProcName As String, _
                                aStoredProcParams() As Variant) As Boolean
                    
    'On Error GoTo ErrorHandler
    Dim cadena As String
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrSrc     As String
    
    Dim oCommand As ADODB.Command
    Dim oParam As ADODB.Parameter
    
    Dim uStoredProcParams() As Variant
    
    Dim iMaxParam As Integer
    Dim iParam As Integer
    Dim cCadena As String
    Dim cNummov As String
    Dim cVoucher As String
    
    cNummov = ""
    cVoucher = ""
    cCadena = ""
    
    uStoredProcParams = aStoredProcParams
    
    'Asumir Falla en Función
    
    ExecuteActionSP = False
    
    ' Instanciamos el Objeto Command
    
    Set oCommand = New ADODB.Command
    oCommand.CommandText = sStoredProcName
    oCommand.CommandType = ADODB.adCmdStoredProc
    
    If gsImportacion = True Then
            oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_LONG_TIMEOUT
    Else
            oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_NORMAL_TIMEOUT
    End If
    
    oCommand.CommandTimeout = 0
    Set oCommand.ActiveConnection = moCachedConnection
    
    ' Si hay Parametros los Metemos al oCommand
    
    iMaxParam = UBound(aStoredProcParams)
    
    If Not (iMaxParam < 0) Then
    
        For iParam = 0 To iMaxParam
        
            ' Creamos el Parametro
           ' iParam = 5
            
            Set oParam = oCommand.CreateParameter( _
                uStoredProcParams(iParam, TEL_SP_PARAMET_NAME), _
                uStoredProcParams(iParam, TEL_SP_PARAMET_TYPE), _
                uStoredProcParams(iParam, TEL_SP_PARAMET_DIRECTION), _
                uStoredProcParams(iParam, TEL_SP_PARAMET_SIZE), _
                uStoredProcParams(iParam, TEL_SP_PARAMET_VALUE))

            ' Lo añadimos a la coleccion de parametros
            
            oCommand.Parameters.Append oParam
            
            
            '----------------------------------------------------------------------------------------
            If CE(uStoredProcParams(iParam, TEL_SP_PARAMET_NAME)) <> "@RETURN_VALUE" Then
                cCadena = cCadena & CE(uStoredProcParams(iParam, TEL_SP_PARAMET_NAME)) & " [" & CE(uStoredProcParams(iParam, TEL_SP_PARAMET_VALUE)) & "], "
            End If

            If UCase(CE(uStoredProcParams(iParam, TEL_SP_PARAMET_NAME))) = "@ASE_CNUMMOV" Then
                cNummov = CE(uStoredProcParams(iParam, TEL_SP_PARAMET_VALUE))
            End If
            
            If UCase(CE(uStoredProcParams(iParam, TEL_SP_PARAMET_NAME))) = "@ASE_NVOUCHER" Then
                cVoucher = CE(uStoredProcParams(iParam, TEL_SP_PARAMET_VALUE))
            End If
            '----------------------------------------------------------------------------------------
            
            
            Set oParam = Nothing
        
        Next
        
    End If
    
    ' Ejecutamos el Comando
    
    If cCadena <> "" Then
        Call EscribirMovimLog("", sStoredProcName, cCadena, cNummov, cVoucher)
    End If
    
'    If gsFilasAfectadas < 0 Then
'        ExecuteActionSP = False
''        GoTo ErrorHandler
'    End If
    oCommand.Execute gsFilasAfectadas
    
    ' Leemos los parametros OUTPUT del Stored Procedure
    
    If Not (iMaxParam < 0) Then
        cadena = "Verifique el siguiente registro ... " & Salto(2)
        For iParam = 0 To iMaxParam
            
            ' Cargamos los parametros de vuelta
            uStoredProcParams(iParam, TEL_SP_PARAMET_SIZE) = _
                                oCommand.Parameters.item(iParam).Size
            uStoredProcParams(iParam, TEL_SP_PARAMET_VALUE) = _
                                oCommand.Parameters.item(iParam).Value
                                
        Next
    End If
    
    Set oCommand = Nothing
    Set oParam = Nothing

    'Indicar Acción de Función Satisfactoria
    
    ExecuteActionSP = True
    Exit Function
    
'ErrorHandler:
'
'            ' Valores de Retorno
'
'            lErrNo = Err.Number
'            sErrDesc = UCase(Err.Description)
'            sErrSrc = Err.Source
'
'            ExecuteActionSP = False
'
'            'On Error Resume Next
'
'            Set oCommand = Nothing
'            Set oParam = Nothing
'
'            On Error GoTo 0
'            Dim msgErr As String
'            Dim PosIni As Integer
'            Dim PosFin As Integer
'
'            Select Case lErrNo
'                Case -2147217873
'
'                        If InStr(1, sErrDesc, "PRIMARY KEY") > 0 Then
'                            msgErr = "Esta intentado insertar un registro que ya EXISTE." & Chr(13)
'                        ElseIf InStr(1, sErrDesc, "INSERT") > 0 Then
'                            If InStr(1, sErrDesc, "NULL") > 0 Then
'                                msgErr = "Esta intentado insertar un valor NULO en un campo obligatorio."
'                            ElseIf InStr(1, sErrDesc, "FOREIGN KEY") > 0 Then
'                                PosIni = InStr(1, sErrDesc, "FK")
'                                PosFin = InStr(PosIni + 1, sErrDesc, "EL CONFLICTO HA APARECIDO")
'                                If (PosFin + 1 - PosIni) < 0 Then
'                                    cadena = ""
'                                Else
'                                    cadena = Mid(sErrDesc, PosIni, PosFin + 1 - PosIni)
'                                End If
'
'                                msgErr = "Esta intentado insertar un valor que no se encuentra la tabla maestra." & Salto(2) & "Mensaje : " & cadena
'                            Else
'                                msgErr = "ERROR AL TRATAR DE INSERTAR UN REGISTRO"
'                            End If
'                        ElseIf InStr(1, sErrDesc, "UPDATE") > 0 Then
'                            If InStr(1, sErrDesc, "NULL") > 0 Then
'                                msgErr = "Esta intentado insertar un valor NULO en un campo obligatorio."
'                            ElseIf InStr(1, sErrDesc, "FOREIGN KEY") > 0 Then
'                                PosIni = InStr(1, sErrDesc, "'FK")
'                                PosFin = InStr(PosIni + 1, sErrDesc, "'")
'                                cadena = Mid(sErrDesc, PosIni, PosFin + 1 - PosIni)
'
'                                msgErr = "Esta intentado actualizar con un valor que no se encuentra la tabla maestra." & Salto(2) & "Mensaje : " & cadena
'                            Else
'                                msgErr = "ERROR AL TRATAR DE ACTUALIZAR EL REGISTRO"
'                            End If
'                        ElseIf InStr(1, sErrDesc, "DELETE") > 0 Then
'                            If InStr(1, sErrDesc, "REFERENCE") > 0 Then
'                                msgErr = "Esta intentado borrar un registro del cual dependen otros datos."
'                            Else
'                                msgErr = "ERROR AL TRATAR DE ELIMINAR EL REGISTRO"
'                            End If
'                        Else
'                            msgErr = Err.Description
'                        End If
'
'                Case Else
'                        msgErr = LCase(sErrDesc)
'            End Select
'
'            Mensajes msgErr, vbInformation + vbOKOnly
''            Resume
End Function

Private Function AsignaValor(nParametro As Integer, valorParametro As String) As String
    Select Case nParametro
           Case 6: AsignaValor = "Libro: " & valorParametro & Salto(1)
           Case 7: AsignaValor = "Voucher: " & valorParametro & Salto(1)
           Case 9: AsignaValor = "Item: " & valorParametro & Salto(1)
           Case 10: AsignaValor = "Cuenta: " & valorParametro & Salto(1)
           Case 16: AsignaValor = "C.Costo: " & valorParametro & Salto(1)
           Case 18: AsignaValor = "Entidad: " & valorParametro & Salto(1)
           Case 19: AsignaValor = "Tipo: " & valorParametro & Salto(1)
           Case 20: AsignaValor = "Serie: " & valorParametro & Salto(1)
           Case 21: AsignaValor = "Num: " & valorParametro & Salto(1)
    End Select
End Function

'**************************************************************************
' Nombre....: GetRecordsetSP()
' Proposito.: Obtiene un Recordset desconectado ejecutando un sp
' Inputs....:
' Se asume..: N/A
' Retorno...: Retorna True si la función tiene exito
' Efectos...: Retrona en el mismo arreglo de parametros los valores
'             de los parametros del SP definidos como OUTPUT.
'**************************************************************************
' Cambio Mensaje de Error
' Optional blnMensaje As Boolean = True

Public Function GetRecordSetSP(ByVal sStoredProcName As String, _
                                oRecordset As ADODB.Recordset, _
                                aStoredProcParams() As Variant, Optional blnMensaje As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrSrc As String
    
    Dim oCommand As ADODB.Command
    Dim oParam As ADODB.Parameter
    
    Dim iParam As Integer
    Dim iMaxParam As Integer
    Dim cCadena As String
    Dim cNummov As String
    Dim cVoucher As String
    
    cNummov = ""
    cVoucher = ""
    cCadena = ""
    
    ' Asumir Falla en la Función
    
    GetRecordSetSP = False
    
    ' Instanciamos el Objeto command
    
    Set oCommand = New ADODB.Command
    
    oCommand.CommandText = sStoredProcName
    oCommand.CommandType = adCmdStoredProc
    Set oCommand.ActiveConnection = moCachedConnection
    
    ' Si hay Parametros los Metemos al oCommand
    
    iMaxParam = UBound(aStoredProcParams)
    
    
    If Not (iMaxParam < 0) Then
    
        For iParam = 0 To iMaxParam
        
            ' Creamos el Parametro
            
            Set oParam = oCommand.CreateParameter( _
                aStoredProcParams(iParam, TEL_SP_PARAMET_NAME), _
                aStoredProcParams(iParam, TEL_SP_PARAMET_TYPE), _
                aStoredProcParams(iParam, TEL_SP_PARAMET_DIRECTION), _
                aStoredProcParams(iParam, TEL_SP_PARAMET_SIZE), _
                aStoredProcParams(iParam, TEL_SP_PARAMET_VALUE))
            
            ' Lo añadimos a la Colección de Parametros
            
            oCommand.Parameters.Append oParam
            
            '----------------------------------------------------------------------------------------
            If CE(aStoredProcParams(iParam, TEL_SP_PARAMET_NAME)) = "@CadenaSQL" Then
                cCadena = cCadena & CE(aStoredProcParams(iParam, TEL_SP_PARAMET_NAME)) & " [" & CE(aStoredProcParams(iParam, TEL_SP_PARAMET_VALUE)) & "], "
            End If


            '----------------------------------------------------------------------------------------
            ' Listo para el siguiente item
            
            Set oParam = Nothing
            
        Next
        
    End If
    
    'Set oRecordset = New ADODB.Recordset
    
    oRecordset.CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
    oRecordset.CursorType = TEL_CURSOR_TYPE.TEL_TYPE_STATIC
    oRecordset.LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
    
    ' Ejecutamos el command, Abrimos el Recordset
    
    
    If gsImportacion = True Then
        oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_LONG_TIMEOUT
    Else
        oCommand.CommandTimeout = TEL_CONNECT_TIMERS_ENUM.TEL_CONNECT_NORMAL_TIMEOUT
    End If
    
    oCommand.CommandTimeout = 0
    oRecordset.Open oCommand
    
    If cCadena <> "" Then
        Call EscribirConsultaLog("", sStoredProcName, cCadena, cNummov, cVoucher)
    End If
    
    ' Leemos los parametros OUTPUT del Stored Procedure
    
    If Not (iMaxParam < 0) Then
    
        For iParam = 0 To iMaxParam
            
            ' Cargamos los parametros de vuelta
            aStoredProcParams(iParam, TEL_SP_PARAMET_SIZE) = _
                                oCommand.Parameters.item(iParam).Size
            aStoredProcParams(iParam, TEL_SP_PARAMET_VALUE) = _
                                oCommand.Parameters.item(iParam).Value
        Next
        
    End If
    
    Set oRecordset.ActiveConnection = Nothing
        'MsgBox oRecordset.Source
    
    Set oCommand = Nothing
    Set oParam = Nothing
'    Set oRecordset = Nothing
    
    GetRecordSetSP = True
    
    Exit Function
    
ErrorHandler:
            'Almacenar Valores de Retorno
            'If Err.Number = 3709 Then
            '    MsgBox "No se pudo establecer la Coneccion, Intentelo nuevamente ", vbInformation, "ECB"
            '    'Err.Clear
            '    'Exit Function
            'End If

            lErrNo = Err.Number
            sErrDesc = Err.Description
            sErrSrc = Err.Source
            
            Set oRecordset = Nothing
            Set oCommand = Nothing
            Set oParam = Nothing
' Cambio Mensaje de Error
        If blnMensaje Then
            'MsgBox "Numero de Error: " & CStr(lErrNo) & vbCrLf & "Modulo: " & msOBJECT_NAME & ".GetRecordsetSP()" & vbCrLf & sErrSrc & vbCrLf & sErrDesc, vbCritical, "Error Grave"
            If lErrNo = 3709 Then
                MsgBox "No se puede establecer la conexión, intentelo nuevamente ", vbInformation, gsNombreModulo
                gsError = True
            Else
                Mensajes "Numero de Error: " & CStr(lErrNo) & vbCrLf & "Modulo: " & msOBJECT_NAME & ".ExecuteActionSP()" & vbCrLf & sErrSrc & vbCrLf & sErrDesc, vbInformation
'                Resume
            End If

        End If

End Function
