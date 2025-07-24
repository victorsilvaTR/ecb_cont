VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFuncionesExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function RetonrarListaSP(strProcedimiento As String, varParametros() As Variant) As ADODB.Recordset

    Dim cmmExec As ADODB.Command
    Dim intX    As Integer
    Dim intTipo As DataTypeEnum
    Set RetonrarListaSP = New ADODB.Recordset
    
    On Error GoTo FuncError
    
    Set cmmExec = New ADODB.Command
    Conectar
    cmmExec.ActiveConnection = gcnSistema
    cmmExec.CommandText = strProcedimiento
    cmmExec.CommandType = adCmdStoredProc
    
    For intX = 0 To 6
        If IsNumeric(varParametros(intX, 1)) Then
            cmmExec.Parameters.Append cmmExec.CreateParameter(varParametros(intX, 0), adVarNumeric, adParamInput, , varParametros(intX, 1))
        Else
            cmmExec.Parameters.Append cmmExec.CreateParameter(varParametros(intX, 0), adVarChar, adParamInput, varParametros(intX, 2), Replace(varParametros(intX, 1), "'", ""))
        End If
    Next
    
    Set RetonrarListaSP = cmmExec.Execute
    Desconectar
Fin:
    On Error Resume Next
    Desconectar
    Exit Function

FuncError:
    Mensajes Str(Err.Number) & " - " & Err.Description, vbInformation
    Resume Fin
    
End Function

Public Function Listar_Registros(arr As Variant, St_Procedure As String, num_parametros As Long) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CMD As New ADODB.Command
    Dim i As Long
    '<>
    rs.CursorLocation = adUseClient
    gtxtSQL = "EXEC " & St_Procedure & " "
    
    If num_parametros <> -1 Then ' Si es -1 significa que el store proce. no tiene parametros
    
            For i = 0 To num_parametros
               If arr(i) = VACIO Then arr(i) = "Null"
               If arr(i) = "''" Then arr(i) = "'Null'"
               
               If arr(i) = VACIO Then
                    arr(i) = "Null"
               End If
               If i = num_parametros Then
                  gtxtSQL = gtxtSQL & arr(i)
               Else
                 gtxtSQL = gtxtSQL & arr(i) & ","
               End If
            Next
     
     End If
     
    gtxtSQL = Trim(gtxtSQL)
    
    With CMD
        .ActiveConnection = gsCadenaConexion
        .CommandType = adCmdText
        .CommandText = gtxtSQL
        .Execute
    End With
    rs.Open CMD, , adOpenDynamic, adLockOptimistic
    Set Listar_Registros = rs
    Set rs = Nothing
    Set CMD = Nothing

End Function

Public Sub Mant_Tablas(arr As Variant, St_Procedure As String, num_parametros As Long)
    On Error GoTo MIERROR
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim i As Long
    '<>
    rs.CursorLocation = adUseClient
    
    If Left(Trim(St_Procedure), 2) = "sp" Or Left(Trim(St_Procedure), 2) = "Sp" Or Left(Trim(St_Procedure), 2) = "sP" Or Left(Trim(St_Procedure), 2) = "SP" Then
    gtxtSQL = "EXEC " & St_Procedure & " "
    Else
    gtxtSQL = St_Procedure & " "
    End If
    
    For i = 0 To num_parametros
       If arr(i) = VACIO Then arr(i) = "Null"
       If arr(i) = "''" Then arr(i) = "Null"
       If arr(i) = "' '" Then arr(i) = "Null"
      If IsNull(arr(i)) = True Then arr(i) = "Null"
       
       If i = num_parametros Then
          gtxtSQL = gtxtSQL & arr(i)
       Else
          gtxtSQL = gtxtSQL & arr(i) & ","
       End If
    Next
    gtxtSQL = Trim(gtxtSQL)
    With CMD
        .ActiveConnection = gsCadenaConexion
        .CommandType = adCmdText
        .CommandText = gtxtSQL
        .Execute
    End With
    Set CMD = Nothing
    Exit Sub
MIERROR:
    MsgBox Err.Description, 16
    
End Sub
Public Function fRetornaValor(cadena As String) As String
    On Error Resume Next
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
      With CMD
        .ActiveConnection = gsCadenaConexion
        .CommandType = adCmdText
        .CommandText = cadena
        .Execute
      End With
     rs.Open CMD, , adOpenDynamic, adLockBatchOptimistic
     If IsNull(rs(0)) = True Or Str(rs(0)) = "" Then
       If Len(Trim(rs(0))) <= 0 Then
         fRetornaValor = ""
        Else
         fRetornaValor = rs(0)
        End If
     Else
        fRetornaValor = rs(0)
     End If
    rs.Close
    Set rs = Nothing
    Set CMD = Nothing
End Function

Public Function fRetornaRS(ByVal cadena_SQL As String) As ADODB.Recordset
    Dim CMD As New ADODB.Command
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
        
    With CMD
      .ActiveConnection = gsCadenaConexion
      .CommandType = adCmdText
      .CommandText = cadena_SQL
      .CommandTimeout = 0
      .Execute
    End With
      
    rs.Open CMD, , adOpenStatic, adLockOptimistic
    
    Set fRetornaRS = rs
    Set rs = Nothing
    Set CMD = Nothing
End Function

Public Sub pEjecutaSQL(cadena_SQL As String)
    Dim CMD As New ADODB.Command
    With CMD
      .ActiveConnection = gsCadenaConexion
      .CommandType = adCmdText
      .CommandText = cadena_SQL
      .CommandTimeout = 0
      .Execute
      End With
    Set CMD = Nothing
End Sub


