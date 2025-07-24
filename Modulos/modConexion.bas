Attribute VB_Name = "modConexion"
Option Explicit
Dim objActivacion As ECActivacionValidacion

Public Function GetRsRecordCount(ByRef rsParam As Recordset) As Long
On Error GoTo errHand
    GetRsRecordCount = 0
    If rsParam Is Nothing Then Exit Function
    If rsParam.State = 0 Then Exit Function
    GetRsRecordCount = rsParam.RecordCount
Exit Function
errHand:
    On Error Resume Next
    If rsParam.BOF = False And rsParam.EOF Then GetRsRecordCount = 1
End Function
Public Sub Conectar()
    On Local Error GoTo ErrorConexion
    If gcnSistema.State = 0 Then
        gcnSistema.ConnectionString = gsCadenaConexion
        gcnSistema.CursorLocation = adUseClient
        gcnSistema.Open
    End If
    Exit Sub
ErrorConexion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub
Public Sub ConectarAdvance()
    On Local Error GoTo ErrorConexion
    If gcnSistemaAdv.State = 0 Then
        Dim sSql As String
        
        sSql = "driver={SQL Server};" & _
                    "server=" & Trim$(gsServidor) & ";uid=" & Trim$(gsBDUS) & ";pwd=" & Trim$(gsBDPW) & _
                    ";database=" & Trim$(gsBD)
        
        gcnSistemaAdv.ConnectionString = sSql
                    
        gcnSistemaAdv.CursorLocation = adUseClient
        gcnSistemaAdv.CommandTimeout = 0
        gcnSistemaAdv.Open

    End If
    Exit Sub
ErrorConexion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
'    Resume
End Sub

Public Sub Desconectar()
    If gcnSistema.State = 1 Then gcnSistema.Close
    Set gcnSistema = Nothing
End Sub

Public Sub CerrarRecordSet(ByRef rsParam As Recordset)
On Error GoTo errHand
    If rsParam Is Nothing Then Exit Sub
    If rsParam.State <> 0 Then rsParam.Close
    Set rsParam = Nothing
Exit Sub
errHand:
    On Error Resume Next
    Set rsParam = Nothing
End Sub

Private Function BuscaParametro(Linea As String) As String
    Dim posicion As Integer
    posicion = InStr(1, Linea, "=")
    BuscaParametro = Trim(Left(Linea, posicion - 1))
End Function

Private Function RetornaValor(Linea As String) As String
    Dim posicion As Integer
    posicion = InStr(1, Linea, "=")
    RetornaValor = Trim(Mid(Linea, posicion + 1, Len(Linea) - posicion))
End Function

Private Function LeeConfiguracion() As Integer
    On Error GoTo serror
    Dim cArchivo As String
    cArchivo = "config.ini"
    LeeConfiguracion = 0
    Dim fso As New Scripting.filesystemobject
    If Not fso.FileExists(App.Path & "\" & cArchivo) Then
        MsgBox "No se encontro archivo de configuracion " & cArchivo & ", copielo y vuelva a ejecutar el sistema", vbInformation, "ERROR..."
    Else
        Dim Linea As String, Opciones As Integer
        Opciones = 0
        
        Open App.Path & "\" & cArchivo For Input As #1
        While Not EOF(1)
            Line Input #1, Linea
            Select Case BuscaParametro(Linea)
                Case "SERVIDOR"
                     gsServidor = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "BASE_DATOS"
                     gsBD = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "AUTENTICACION_WINDOWS"
                     gsAutenticacion = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "USUARIO"
                     gsBDUS = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "PWD"
                     gsBDPW = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "CONTRASEÑA"
                     gsBDPW = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "BACKUP"
                     gsRutaBackup = RetornaValor(Linea)
                     Opciones = Opciones + 1
            End Select
        Wend
        Close #1
    End If
    Set fso = Nothing

    If gsEncriptada = "SI" Then
        gsBDPW = DesCifrar(gsBDPW, "977611")
    End If
    LeeConfiguracion = Opciones
    Exit Function
serror:
    LeeConfiguracion = 0
End Function

Private Function LeeConfigAdic() As Integer
    On Error GoTo serror
    Dim cArchivo As String
    cArchivo = "configadic.ini"
    LeeConfigAdic = 0
    Dim fso As New Scripting.filesystemobject
    If Not fso.FileExists(App.Path & "\" & cArchivo) Then
        MsgBox "No se encontro archivo de configuracion " & cArchivo & ", copielo y vuelva a ejecutar el sistema", vbInformation, "ERROR..."
    Else
        Dim Linea As String, Opciones As Integer
        Opciones = 0
        
        Open App.Path & "\" & cArchivo For Input As #1
        While Not EOF(1)
            Line Input #1, Linea
            Select Case BuscaParametro(Linea)
                Case "GEN_DSN"
                     gsGenDsn = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "ENCRIPTADA"
                     gsEncriptada = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "GENLOG"
                     gsGenLog = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "GENLOGMOV"
                     gsGenLogMov = RetornaValor(Linea)
                     Opciones = Opciones + 1
                Case "GENLOGLEC"
                     gsGenLogConsulta = RetornaValor(Linea)
                     Opciones = Opciones + 1
            End Select
        Wend
        Close #1
    End If
    Set fso = Nothing
     Set fso = New clsCommonFuncs
     
    LeeConfigAdic = Opciones
    Exit Function
serror:
    LeeConfigAdic = 0
End Function

Public Sub Main()

    Set objActivacion = New ECActivacionValidacion
    gsCadenaConexion = ""
    
    Call LeeConfigAdic
    
    If LeeConfiguracion < 6 Then
        MsgBox "Verifique el archivo de configuracion, faltan parametros iniciales", vbInformation, "Error"
    Else
    
        If UCase(gsAutenticacion) = "TRUE" Then
            gsCadenaConexion = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & gsBD & ";Data Source=" & gsServidor
        Else
            gsCadenaConexion = "Provider=SQLOLEDB.1;password = " & gsBDPW & ";User ID=" & gsBDUS & ";Initial Catalog=" & gsBD & ";Data Source=" & gsServidor
        End If
        
        Dim cn As New ADODB.Connection
        cn.Open gsCadenaConexion
        
        If cn.State = 1 Then cn.Close
    
    
    objActivacion.Main
    
    ''PGBV:Este segmento se comenta y solo se invoca a frmPrcIngresoSistema.show -- Univ. Pacifico
     '   If ValidaRegistro = False Then
     '       frmRegistro.Show
     '   Else
     '       objActivacion.Main
     '       'frmPrcIngresoSistema.Show
            
     '   End If



'If ValidaRegistroEcb = False Then
'frmRegistro.Show
'Else
'frmPrcIngresoSistema.Show
'End If



    End If
    Call SaberVersionWindows
End Sub
Private Function Activa_ParamIniciales()
On Error GoTo Control
Exit Function
Control:
 MsgBox Err.Description
End Function
