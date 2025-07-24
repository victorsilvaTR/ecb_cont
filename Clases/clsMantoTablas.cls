VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMantoTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CadenaConexion As String
Private m_NombreSP As String
Private m_FlagCnx As Boolean
Private m_EjecutarConsulta As Boolean
Private objDatos As IDataAccess
Private m_Tabla As ADODB.Recordset

Public Property Get CadenaDeConexion() As String
    CadenaDeConexion = m_CadenaConexion

End Property

Public Property Let CadenaDeConexion(ByVal strData As String)
    m_CadenaConexion = strData
End Property

Public Property Get NombreSP() As String
    NombreSP = m_NombreSP
End Property

Public Property Let NombreSP(ByVal strData As String)
    m_NombreSP = strData
End Property

Private Sub Class_Initialize()
    m_FlagCnx = False
End Sub

Private Sub CerrarRecordSet(ByRef rsParam As Recordset)
    If rsParam Is Nothing Then Exit Sub
    If rsParam.State <> 0 Then rsParam.Close
    Set rsParam = Nothing
End Sub

Private Sub Class_Terminate()
    CerrarRecordSet m_Tabla
    If Not objDatos Is Nothing Then
        objDatos.DestroyConecction
    End If
End Sub

Public Function InicializaClase()
    Set m_Tabla = New ADODB.Recordset
    Set objDatos = New IDataAccess
    objDatos.CacheConnection (CadenaDeConexion)
End Function

Public Function BeginTrans()
    objDatos.IniciarTransaccion
End Function

Public Function CommitTrans()
    objDatos.FinalizarTransaccion
End Function

Public Function CancelTrans()
    objDatos.CancelarTransaccion
End Function

Public Function FinalizaClase()
    objDatos.DestroyConecction
End Function

Public Function MantenimientoDeTablas(ByVal strCnx As String, ByVal StoredProcedure As String, Parametros(), _
                                      Optional ByVal blnStatus As Boolean, _
                                      Optional ByVal bForzarCommit As Boolean = False, _
                                      Optional ByVal bSinTransacciones As Boolean = False) As Boolean
    
    '--- blnStatus : fuerza a cerrar la conexion si se termina la transaccion
    
    CadenaDeConexion = strCnx
    
    If Not m_FlagCnx Then
        InicializaClase
        
        If bSinTransacciones = False Then
            Me.BeginTrans
        End If
        
        m_FlagCnx = True
    End If
    
    NombreSP = StoredProcedure
    m_EjecutarConsulta = False
    
    Call ObtenerEstructuraSP
    
    If Not m_Tabla Is Nothing Then
        If m_Tabla.RecordCount > 0 Then
            MantenimientoDeTablas = EjecutarTransaccion(Parametros)
        Else
            MantenimientoDeTablas = False
        End If
    End If
    
    If bSinTransacciones = False Then
        If MantenimientoDeTablas = True Then
            If blnStatus Then ' si se ejecuto la transaccion y el blnstatus = True entonces fuerza a cerrar la clase
                Me.CommitTrans
                Me.FinalizaClase
                m_FlagCnx = False
            End If
        Else
            If bForzarCommit = True Then
                Me.CommitTrans 'si sale un error igual hacer un commit a la transaccion
            Else
                Me.CancelTrans 'si sale un error hace un rollback de la transaccion
            End If
            
            If blnStatus = True Then
                Me.FinalizaClase
                m_FlagCnx = False
            End If
        End If
    Else
        If blnStatus = True Then
            Me.FinalizaClase
            m_FlagCnx = False
        End If
    End If
    
End Function

Public Function ConsultaDatosTabla(ByVal strCnx As String, ByVal StoredProcedure As String, Parametros(), Optional blnMensaje As Boolean = True) As ADODB.Recordset
    CadenaDeConexion = strCnx
    InicializaClase
    NombreSP = StoredProcedure
    m_EjecutarConsulta = True
    
    ObtenerEstructuraSP
    
    If Not m_Tabla Is Nothing Then
        If m_Tabla.RecordCount > 0 Then
            If EjecutarTransaccion(Parametros, blnMensaje) Then
                Set ConsultaDatosTabla = m_Tabla.Clone
            Else
                Set ConsultaDatosTabla = Nothing
            End If
        End If
    End If

    FinalizaClase
End Function

Private Function EjecutarTransaccion(Parametros(), Optional blnMensaje As Boolean = True) As Boolean
    
    On Error GoTo ErrHandler
    Dim intNroFil As Integer
    Dim intDireccion As Integer
    Dim lngNroFilas As Long
    lngNroFilas = m_Tabla.RecordCount
    
    Dim arrParamCmd() As Variant
    ReDim arrParamCmd(lngNroFilas, 4) As Variant
    Dim TmpintNroFil As Integer
    Dim sw As Integer
    sw = 0
    
    
    arrParamCmd(0, TEL_SP_PARAMET_NAME) = "@RETURN_VALUE"
    arrParamCmd(0, TEL_SP_PARAMET_TYPE) = adInteger
    arrParamCmd(0, TEL_SP_PARAMET_DIRECTION) = TEL_PARAM_RETURN_VALUE
    arrParamCmd(0, TEL_SP_PARAMET_SIZE) = 0
    arrParamCmd(0, TEL_SP_PARAMET_VALUE) = vbEmpty
    intNroFil = 1
    
    If NombreSP = "spCn_GeneraAsientoEjer_Simple" Then
        'Erase arrParamCmd
       ' ReDim arrParamCmd(lngNroFilas, 4) As Variant
        'intNroFil = 0
        sw = 1
        TmpintNroFil = 1
    End If
    
    m_Tabla.MoveFirst
    
    Do While Not m_Tabla.EOF
        
        ' Nombre del Parámetro
        arrParamCmd(intNroFil, TEL_SP_PARAMET_NAME) = m_Tabla!Nombre
        
        ' Tipo de Dato VBasic
        If UCase$(m_Tabla!Tipo) = "DATETIME" Then
            arrParamCmd(intNroFil, TEL_SP_PARAMET_TYPE) = adDate
        Else
            arrParamCmd(intNroFil, TEL_SP_PARAMET_TYPE) = fTipoDatoVB("ad" + StrConv(m_Tabla!Tipo, vbProperCase))
        End If
        
        ' Direccion del Parámetro
        If m_Tabla!ISOUTPARAM = 0 Then
            intDireccion = TEL_PARAM_INPUT
        Else
            intDireccion = TEL_PARAM_INPUT_OUTPUT
        End If
        
        arrParamCmd(intNroFil, TEL_SP_PARAMET_DIRECTION) = intDireccion
        
        ' Longitud
        arrParamCmd(intNroFil, TEL_SP_PARAMET_SIZE) = m_Tabla!Longitud
    
        ' Valor a Pasar al Parámetro
        
        If sw = 1 Then TmpintNroFil = TmpintNroFil + 1 Else TmpintNroFil = intNroFil
        If TmpintNroFil - 1 > UBound(Parametros) Then
            arrParamCmd(intNroFil, TEL_SP_PARAMET_VALUE) = vbNull
        Else
            arrParamCmd(intNroFil, TEL_SP_PARAMET_VALUE) = Parametros(TmpintNroFil - 1)
        End If
        
        m_Tabla.MoveNext
        intNroFil = intNroFil + 1
    Loop


    Dim blnResultado As Boolean
        'm_Tabla.MoveFirst
        
    If Not m_EjecutarConsulta Then
        m_Tabla.Close
        blnResultado = objDatos.ExecuteActionSP(NombreSP, arrParamCmd)
        If blnResultado = True Then
            Dim i As Integer
            For i = 1 To UBound(arrParamCmd)
                If i - 1 <= UBound(Parametros) Then
                    Parametros(i - 1) = arrParamCmd(i, 4)
                End If
            Next i
            EjecutarTransaccion = True
        Else
            EjecutarTransaccion = False
        End If
    Else
        m_Tabla.Close
        objDatos.GetRecordSetSP NombreSP, m_Tabla, arrParamCmd, blnMensaje
        If m_Tabla.RecordCount > 0 Then
            EjecutarTransaccion = True
        Else
            EjecutarTransaccion = False
        End If
    End If
sw = 0
    Exit Function

ErrHandler:

    If blnMensaje Then
        Mensajes "Error al ejecutar la transaccion, intente nuevamente..." & Salto(2) & Err.Source & Salto(1) & Err.Description, vbInformation
    End If
    EjecutarTransaccion = False
End Function
    
Private Function ObtenerEstructuraSP() As Boolean
    On Error GoTo ErrHandler
    Dim arrParamCmd(1, 4) As Variant
        
    arrParamCmd(0, TEL_SP_PARAMET_NAME) = "@RETURN_VALUE"
    arrParamCmd(0, TEL_SP_PARAMET_TYPE) = adInteger
    arrParamCmd(0, TEL_SP_PARAMET_DIRECTION) = TEL_PARAM_RETURN_VALUE
    arrParamCmd(0, TEL_SP_PARAMET_SIZE) = 0
    arrParamCmd(0, TEL_SP_PARAMET_VALUE) = vbEmpty
    
    arrParamCmd(1, TEL_SP_PARAMET_NAME) = "@strNombreSP"
    arrParamCmd(1, TEL_SP_PARAMET_TYPE) = adVarChar
    arrParamCmd(1, TEL_SP_PARAMET_DIRECTION) = TEL_PARAM_INPUT
    arrParamCmd(1, TEL_SP_PARAMET_SIZE) = 30
    arrParamCmd(1, TEL_SP_PARAMET_VALUE) = NombreSP
    
    objDatos.GetRecordSetSP "usp_EstructuraSP", m_Tabla, arrParamCmd
    
    ObtenerEstructuraSP = True
    Exit Function
ErrHandler:
    'MsgBox "Error Nº " & CStr(Err.Number) & vbCrLf & "IStdMTSClass" & ".ObtenerEstructuraSP()" & vbCrLf & Err.Source & vbCrLf & Err.Description, vbCritical, "Error de Ejecución"
    Mensajes "Error Nº " & CStr(Err.Number) & vbCrLf & "IStdMTSClass" & ".ObtenerEstructuraSP()" & vbCrLf & Err.Source & vbCrLf & Err.Description, vbInformation
    ObtenerEstructuraSP = False
End Function

Private Function fTipoDatoVB(ByVal strTipoDato) As Integer
    ' *** Convierte el tipo de dato sql a tipo de dato Basic
    Select Case strTipoDato
        Case "adInt"        ' *** Parche Nº 01 - 09/09/2004 - Johnny Huaman Mendoza
            fTipoDatoVB = adInteger
        Case "adInteger"
            fTipoDatoVB = adInteger
        Case "adVarchar"
            fTipoDatoVB = adVarChar
        Case "adChar"
            fTipoDatoVB = adChar
        Case "adBoolean"
            fTipoDatoVB = adBoolean
        Case "adBinary"
            fTipoDatoVB = adBinary
        Case "adDecimal"
            fTipoDatoVB = adDouble
        Case "adDouble"
            fTipoDatoVB = adDouble
        Case "adFloat"
            fTipoDatoVB = adDouble
        Case "adNumeric"
            fTipoDatoVB = adDouble
    End Select
End Function



