Attribute VB_Name = "modArreglos"
Option Explicit


Public Sub LLenarArregloRS(ByRef arreglo As XArrayDB, ByRef rsArreglo As ADODB.Recordset)
    Dim i As Integer, j As Integer
    Dim nCOLS As Integer, nROWS As Integer
    If rsArreglo Is Nothing Then Exit Sub
    
    nCOLS = rsArreglo.Fields.Count - 1
    nROWS = GetRsRecordCount(rsArreglo)
    arreglo.ReDim 0, nROWS - 1, 0, nCOLS
    ' *** Llenando el Arreglo
    For i = 0 To nROWS - 1
      For j = 0 To nCOLS
        arreglo(i, j) = rsArreglo(rsArreglo.Fields(j).Name).Value
      Next
      rsArreglo.MoveNext
    Next
End Sub


Public Function LlenarArreglo(arreglo As XArrayDB, Sql As String, Optional Filtro As String = "") As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Col As Integer
    Dim rsArreglo As ADODB.Recordset
'''''    Call Conectar
'''''    Set rsArreglo = gcnSistema.Execute(sql)
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    arrDatos = Array(Sql)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Filtro <> "" Then rsArreglo.Filter = Filtro
    
    arreglo.Clear
    arreglo.ReDim 0, 0, 0, 0
    
    If rsArreglo Is Nothing Then
       On Error Resume Next
       Call CerrarRecordSet(rsArreglo)
       Set clDatos = Nothing
       Exit Function
    End If
    
    Col = rsArreglo.Fields.Count - 1
    arreglo.ReDim 0, 0, 0, Col
    i = 0
    ' *** Llenando el Arreglo
    Do While Not rsArreglo.EOF
      If i > 0 Then arreglo.ReDim 0, i + 1, 0, Col
      
      i = i + 1
      For j = 0 To rsArreglo.Fields.Count - 1
        arreglo(i - 1, j) = rsArreglo(rsArreglo.Fields(j).Name).Value
      Next
      rsArreglo.MoveNext
    Loop
    ' *** Redimensionando arreglo y cerrando el recordSet
    arreglo.ReDim 0, i - 1, 0, Col
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    ' ***
    LlenarArreglo = i - 1
End Function

Public Sub LlenarArregloRetornandoFilas(arreglo As XArrayDB, Sql As String, ByRef FilasAfectadas As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim Col As Integer
    Dim rsArreglo As ADODB.Recordset
'''''    Call Conectar
'''''    Set rsArreglo = gcnSistema.Execute(sql)
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    arrDatos = Array(Sql)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    arreglo.Clear
    arreglo.ReDim 0, 0, 0, 0
    
    
    If rsArreglo Is Nothing Then
        FilasAfectadas = 0
        On Error Resume Next
        Call CerrarRecordSet(rsArreglo)
        Set clDatos = Nothing
        Exit Sub
    End If
    
    FilasAfectadas = rsArreglo.RecordCount
    
    Col = rsArreglo.Fields.Count - 1
    arreglo.ReDim 0, 0, 0, Col
    i = 0
    ' *** Llenando el Arreglo
    Do While Not rsArreglo.EOF
      If i > 0 Then arreglo.ReDim 0, i + 1, 0, Col
      
      i = i + 1
      For j = 0 To rsArreglo.Fields.Count - 1
        arreglo(i - 1, j) = rsArreglo(rsArreglo.Fields(j).Name).Value
      Next
      rsArreglo.MoveNext
    Loop
    ' *** Redimensionando arreglo y cerrando el recordSet
    arreglo.ReDim 0, i - 1, 0, Col
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Sub

Public Sub LlenarArregloAdicionandoNinguno(arreglo As XArrayDB, Sql As String, Texto As String)
    Dim i As Integer
    Dim j As Integer
    Dim Col As Integer
    Dim rsArreglo As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    arrDatos = Array(Sql)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    arreglo.Clear
    arreglo.ReDim 0, 0, 0, 0
    
    If rsArreglo Is Nothing Then
       On Error Resume Next
       Call CerrarRecordSet(rsArreglo)
       Set clDatos = Nothing
       Exit Sub
    End If
    
    Col = rsArreglo.Fields.Count - 1
    arreglo.ReDim 0, 0, 0, Col
    i = 0
    
    ' ---------- Llenando el Arreglo -----------
    arreglo.ReDim 0, i + 1, 0, Col
    arreglo(0, 1) = Texto
    i = i + 1
    '-------------------------------------------
    Do While Not rsArreglo.EOF
      If i > 1 Then arreglo.ReDim 0, i + 1, 0, Col
      i = i + 1
      For j = 0 To rsArreglo.Fields.Count - 1
        arreglo(i - 1, j) = rsArreglo(rsArreglo.Fields(j).Name).Value
      Next
      rsArreglo.MoveNext
    Loop

    arreglo.ReDim 0, i - 1, 0, Col
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing

End Sub
Public Sub ComboArregloBaseImp(arr As XArrayDB, combo As TDBCombo, cadena As String)
    Call LlenarArregloAdicionandoNinguno(arr, cadena, "")
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    combo.ReBind
End Sub

Public Sub ComboArregloNinguno(arr As XArrayDB, combo As TDBCombo, cadena As String)
    Call LlenarArregloAdicionandoNinguno(arr, cadena, "<NINGUNO>")
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    combo.ReBind
End Sub

Public Sub ComboArreglo(arr As XArrayDB, combo As TDBCombo, cadena As String, Optional Filtro As String = "")
    Call LlenarArreglo(arr, cadena, Filtro)
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    combo.ReBind
End Sub

Public Sub ComboArregloMonAdic(arr As XArrayDB, combo As TDBCombo, cadena As String)
    Call LlenarArregloAdicionandoNinguno(arr, cadena, "")
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column4"
    combo.BoundColumn = "column0"
    combo.ReBind
End Sub

Public Sub LlenarRecordSet(Sql As String, rsTabla As Recordset, Optional mensaje As Boolean = True)
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Set rsTabla = New ADODB.Recordset
    Dim arrDatos() As Variant
    arrDatos = Array(Sql)
    
    If Not rsTabla Is Nothing Then
        If rsTabla.State = 1 Then
            rsTabla.Close
        End If
    End If
    
    Set rsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos(), mensaje)
    Set clDatos = Nothing
End Sub

Public Sub GridArreglo(arr As XArrayDB, grilla As TDBGrid, cadena As String, Optional Filtro As String = "")
    Call LlenarArreglo(arr, cadena, Filtro)
    Set grilla.Array = arr
    grilla.ReBind
End Sub

Public Function BuscaValorArr(arreglo As XArrayDB, Valor As Variant, columna As Integer)
    Dim i As Integer
    BuscaValorArr = -1
    For i = 0 To arreglo.Count(1) - 1
        If arreglo(i, columna) = Valor Then
            BuscaValorArr = i
            Exit For
        End If
    Next
End Function

Public Sub LlenaComboMes(arr As XArrayDB, combo As TDBCombo)
    ' *** Llena el areglo del mes
    Dim i As Integer
    
    arr.ReDim 0, 11, 0, 1
    For i = 0 To 11
        arr(i, 0) = Format(i + 1, "00"):  arr(i, 1) = UCase(MonthName(i + 1))
    Next
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    ' ***
End Sub

Public Sub LlenaComboMesApe(arr As XArrayDB, combo As TDBCombo)
    ' *** Llena el areglo del mes
    Dim i As Integer
    
    arr.ReDim 0, 14, 0, 1
    arr(0, 0) = "00": arr(0, 1) = "APERTURA"
    For i = 0 To 11
        arr(i + 1, 0) = Format(i + 1, "00"): arr(i + 1, 1) = UCase(MonthName(i + 1))
    Next
    arr(13, 0) = "13": arr(13, 1) = "AJUSTE"
    arr(14, 0) = "14": arr(14, 1) = "CIERRE"
    Set combo.Array = arr
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    ' ***
End Sub

Public Sub LlenaComboMesApeAddItem(combo As TDBCombo)
    ' *** Llena el areglo del mes
    Dim i As Integer
    Dim cadena As String
    
    combo.AddItem "00" & ";" & "APERTURA"
    For i = 0 To 11
        combo.AddItem Format(i + 1, "00") & ";" & UCase(MonthName(i + 1))
    Next
    combo.AddItem "13" & ";" & "AJUSTE"
    combo.AddItem "14" & ";" & "CIERRE"
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    ' ***
End Sub

Public Sub LlenaComboMesActivo(combo As TDBCombo)
    ' *** Llena el areglo del mes
    Dim i As Integer
    Dim cadena As String
    
    For i = 0 To 11
        combo.AddItem Format(i + 1, "00") & ";" & UCase(MonthName(i + 1))
    Next
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    ' ***
End Sub

Public Sub LlenaComboMesAddItemIniFin(combo As TDBCombo, sMesIni As String, sMesFin As String)
    ' *** Llena el areglo del mes
    Dim i As Integer
    Dim cadena As String
    
    For i = Val(sMesIni) To Val(sMesFin)
        If i = 0 Then
            combo.AddItem "00;APERTURA"
            
        ElseIf i >= 1 And i <= 12 Then
            combo.AddItem Format(i, "00") & ";" & UCase(MonthName(i))
            
        ElseIf i = 13 Then
            combo.AddItem "13;AJUSTE"
        
        ElseIf i = 14 Then
            combo.AddItem "14;CIERRE"
        End If
        
         
    Next
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    
    ' ***
End Sub

Public Sub LlenaComboMesAddItem(combo As TDBCombo, Optional CampoBlanco As Boolean = False, Optional TextoNinguno As Boolean = False, Optional TextoDescrip As String = "", Optional bLimpiarCombo As Boolean = True, Optional TextoCodigo As String = "")
    ' *** Llena el areglo del mes
    Dim i As Integer
    Dim cadena As String
    Dim sTexto As String
    
    If bLimpiarCombo = True Then combo.Clear
    
    If TextoDescrip = "" Then
       sTexto = "<NINGUNO>"
    Else
       sTexto = TextoDescrip
    End If
    
    If CampoBlanco = True Then
        If TextoNinguno = True Then
           combo.AddItem TextoCodigo + ";" + sTexto
        Else
           combo.AddItem "" + ";" + ""
        End If
    End If
    
    
    For i = 0 To 11
        combo.AddItem Format(i + 1, "00") & ";" & UCase(MonthName(i + 1))
    Next
    
    On Error Resume Next
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    combo.ReBind

End Sub

Public Sub LlenaComboMesAddItemAPE(combo As TDBCombo)
    Dim i As Integer
    Dim cadena As String
    Dim sMes As String
    For i = 0 To 12
        sMes = Right("00" & CStr(i), 2)
        combo.AddItem sMes & ";" & NombreMes(sMes)
    Next
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
 

End Sub

Public Function LlenarComboAddItem(ByRef combo As TDBCombo, Sql As String, Optional CampoBlanco As Boolean = False, Optional TextoNinguno As Boolean = False, Optional TextoDescrip As String = "", Optional bLimpiarCombo As Boolean = True, Optional TextoCodigo As String = "") As Integer
    DoEvents
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim j As Integer
    Dim cadena As String
    Dim sTexto As String
    LlenarComboAddItem = 0
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(Sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If bLimpiarCombo = True Then combo.Clear
    
    If TextoDescrip = "" Then
       sTexto = "<NINGUNO>"
    Else
       sTexto = TextoDescrip
    End If
    
    If CampoBlanco = True Then
        If TextoNinguno = True Then
           combo.AddItem TextoCodigo + ";" + sTexto
        Else
           combo.AddItem "" + ";" + ""
        End If
    End If
    ' ***
    If rsAddItem Is Nothing Then
       On Error Resume Next
       Call CerrarRecordSet(rsAddItem)
       Set clDatos = Nothing
       Exit Function
    End If
    ' *** Llenando el Arreglo
    Do While Not rsAddItem.EOF
      cadena = ""
      For j = 0 To rsAddItem.Fields.Count - 1
        If j = rsAddItem.Fields.Count - 1 Then
            cadena = cadena + CE(rsAddItem(rsAddItem.Fields(j).Name))
        Else
            cadena = cadena + CE(rsAddItem(rsAddItem.Fields(j).Name)) + ";"
        End If
      Next
      On Error Resume Next
      combo.AddItem cadena
      
      rsAddItem.MoveNext
    Loop
    LlenarComboAddItem = rsAddItem.RecordCount
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    On Error Resume Next
    combo.Bookmark = 0
    combo.ListField = "column1"
    combo.BoundColumn = "column0"
    combo.ReBind
    
End Function

