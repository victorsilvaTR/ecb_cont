VERSION 5.00
Begin VB.Form FormLoadNew 
   BackColor       =   &H80000009&
   Caption         =   "New Load"
   ClientHeight    =   3156
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6612
   LinkTopic       =   "Form3"
   ScaleHeight     =   3156
   ScaleWidth      =   6612
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   252
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   3072
      Left            =   60
      Picture         =   "FormLoadNew.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   6492
   End
End
Attribute VB_Name = "FormLoadNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONNECTION_STRING As String = "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=ECActivacion;Integrated Security=SSPI;"
Public Function ExisteBD() As Boolean
    On Error GoTo ErrHandler
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    conn.Open CONNECTION_STRING
    
    ExisteBD = True
        conn.Close
    Set conn = Nothing
    Exit Function

ErrHandler:
    ExisteBD = False
End Function
Public Function EstaVaciaTabla() As Boolean
    On Error GoTo ErrHandler
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    conn.Open CONNECTION_STRING
    rs.Open "SELECT TOP 1 Id FROM sfdc_activar_det_sw_ecbcont_peru", conn

    EstaVaciaTabla = True
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

ErrHandler:
    EstaVaciaTabla = True ' Considerar vacía si hay error
End Function
Public Function HayEstadoCritico() As Boolean
    On Error GoTo ErrHandler
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    conn.Open CONNECTION_STRING
    rs.Open "SELECT TOP 1 Id FROM sfdc_activar_det_sw_ecbcont_peru WHERE Estado_Id IN (6, 7)", conn

    HayEstadoCritico = Not rs.EOF
    
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

ErrHandler:
    HayEstadoCritico = False
End Function
Public Sub EjecutarLogica()
    Dim json As Object
    Dim jsonText As String

    If ExisteBD() Then
        If EstaVaciaTabla() Then
            MsgBox "La base existe y está vacía. Llamando al servicio para actualizar...", vbInformation
            
            jsonText = GetDataService()
            If jsonText = "" Then
                MsgBox "Error al obtener datos de la BD"
                Exit Sub
            End If

            Set json = JsonConverter.ParseJson(jsonText)

            ' Si necesitas una clase act con esos datos:
            Dim act As clsActivacion
            Set act = ObtenerObjetoDesdeJson(json)
            
            If Not act Is Nothing Then
                Call ReemplazarRegistroActivacion(act)
            End If
        End If

        If HayEstadoCritico() Then
            MsgBox "La base está en estado crítico", vbExclamation
            Shell "notepad.exe", vbNormalFocus
        End If
    Else
        MsgBox "La base no existe", vbExclamation
        Shell "notepad.exe", vbNormalFocus
    End If
End Sub
Public Function ObtenerObjetoDesdeJson(json As Object) As clsActivacion
    Dim act As New clsActivacion
    On Error GoTo Err
    act.Id = json("ID")
    act.Opportunity_ID = json("OP")
    act.Tax_Number = json("TAX_NUMBER")
    act.Product_Code = json("PRODUCT_CODE")
    act.Product_Name = json("PRODUCT_NAME")
    act.FechaInicio = json("FECHA_INI")
    act.FechaFin = json("FECHA_FIN")
    act.Fecha_Activacion = json("FECHA_ACTIVACION")
    act.Fecha_Desactivacion = "" & json("FECHA_DESACTIVACION")
    
    act.Codigo1 = json("CODIGO1")
    act.Codigo2 = json("CODIGO2")
    act.Host = json("HOST")
    act.Sist_op_desc = json("SISOP")
    act.Sist_op_version = json("SISVER")
    act.Suscribe = json("SUSCRIBE")
    act.Estado_Id = json("ESTADO_ID")
    Set ObtenerObjetoDesdeJson = act
    Exit Function
Err:
    Set ObtenerObjetoDesdeJson = Nothing
End Function
Private Sub Command1_Click()
Call EjecutarLogica

    Dim Existe As Boolean
    Dim vacia As Boolean
    Dim estadoCritico As Boolean

    Existe = ExisteBD()
    vacia = EstaVaciaTabla()
    estadoCritico = HayEstadoCritico()

    MsgBox "ExisteBD: " & Existe & vbCrLf & _
           "EstaVaciaTabla: " & vacia & vbCrLf & _
           "HayEstadoCritico: " & estadoCritico

End Sub

Function ObtenerDatosConexion(ByRef taxNumber As String, ByRef Host As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim conn As Object
    Dim rs As Object
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    conn.Open CONNECTION_STRING
    
    rs.Open "SELECT TOP 1 * from sfdc_activar_det_sw_ecbcont_peru", conn
    If Not rs.EOF Then
        taxNumber = rs.Fields("Tax_Number").Value
        Host = rs.Fields("Host").Value
        ObtenerDatosConexion = True
    Else
        ObtenerDatosConexion = False
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

ErrHandler:
    MsgBox "Error al obtener datos: " & Err.Description
    ObtenerDatosConexion = False
End Function
Function GetDataService() As String
    On Error GoTo Errores

    ' Parámetros que vendrán desde la base de datos
    Dim taxNumber As String
    Dim Host As String
    Dim tokenRequestId As String
    tokenRequestId = "KlIybTVuMnIxYzM0bjJzV1NwMSRzdzRyZCQuMjAwMF9d"

    ' Obtener los datos desde SQL Server
    If Not ObtenerDatosConexion(taxNumber, Host) Then
        GetDataService = "Error al obtener datos de la BD"
        Exit Function
    End If
  
    ' URL del servicio
    Dim WSHTTPCont As String
    WSHTTPCont = "http://activacion.thomsonreuters.cl/WSECActivacion/Service.asmx"

    ' Crear solicitud SOAP
    Dim HttpRequest As Object
    Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")

    Dim payload As String
    payload = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
              "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
              "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
              " xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
              "<soap:Body>" & _
                "<GetActivationCurrent xmlns=""http://tempuri.org/"">" & _
                  "<tokenRequestId>" & tokenRequestId & "</tokenRequestId>" & _
                  "<taxNumber>" & taxNumber & "</taxNumber>" & _
                  "<host>" & Host & "</host>" & _
                "</GetActivationCurrent>" & _
              "</soap:Body>" & _
              "</soap:Envelope>"

    ' Enviar solicitud
    
    With HttpRequest
        .Open "POST", WSHTTPCont, False
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://tempuri.org/GetActivationCurrent"
        .send payload
    End With

    ' Procesar respuesta
    Dim xmldoc As Object
    Set xmldoc = CreateObject("MSXML2.DOMDocument")
    xmldoc.async = False
    xmldoc.LoadXML HttpRequest.responseText

    Dim node As Object
    Set node = xmldoc.SelectSingleNode("//GetActivationCurrentResult") ' Ajusta este nodo si es necesario
    
    If node Is Nothing Then
    MsgBox "Error: No se encontró el nodo GetActivationCurrentResult."
    ElseIf Trim(node.Text) = "" Or LCase(Trim(node.Text)) = "null" Then
    MsgBox "El servicio respondió con un valor nulo o vacío."
    Else
    
    
    Dim cleanJson As String
    cleanJson = node.Text
' Reemplazar comillas tipográficas por comillas normales
    cleanJson = Replace(cleanJson, "“", """")
    cleanJson = Replace(cleanJson, "”", """")

' También quitar espacios innecesarios al inicio/final
    cleanJson = Trim(cleanJson)
    
    GetDataService = cleanJson
    
    End If
    
    Exit Function

Errores:
    GetDataService = False
End Function

Function ObtenerObjetoActivacionJson() As clsActivacion
    On Error GoTo Errores
    Dim json As Object
    Dim jsonText As String
    Dim obj As New clsActivacion
    jsonText = GetDataService()
    
    If jsonText = "" Then Exit Function
    Set json = JsonConverter.ParseJson(jsonText)
    
    With obj
        .Id = CLng(json("Id"))
        .Estado_Id = CLng(json("Estado_Id"))
        .FechaInicio = CDate(json("FechaInicio"))
        .FechaFin = CDate(json("FechaFin"))
        .Fecha_Activacion = CDate(json("Fecha_Activacion"))
        .Fecha_Desactivacion = CDate(json("Fecha_Desactivacion"))

        .Tax_Number = json("Tax_Number")
        .Opportunity_ID = json("Opportunity_ID")
        .Product_Code = json("Product_Code")
        .Product_Name = json("Product_Name")
        .Codigo1 = json("Codigo1")
        .Codigo2 = json("Codigo2")
        .Codigo_Desact = json("Codigo_Desact")
        .Host = json("Host")
        .Sist_op_desc = json("Sist_op_desc")
        .Sist_op_version = json("Sist_op_version")
        .Suscribe = CBool(json("Suscribe"))
    End With

    ObtenerObjetoActivacionJson = obj
    Exit Function

Errores:
    MsgBox "Error: " & Err.Description
    Set ObtenerObjetoActivacionJson = Nothing
End Function

Public Sub ReemplazarRegistroActivacion(ByVal activacion As clsActivacion)
    On Error GoTo Errores

    Dim conn As ADODB.Connection
    Dim sqlDelete As String
    Dim sqlInsert As String

    ' Abrir conexión
    Set conn = New ADODB.Connection
    conn.Open "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=ECActivacion;Integrated Security=SSPI"

    ' Eliminar todos los registros existentes
    sqlDelete = "DELETE FROM sfdc_activar_det_sw_ecbcont_peru"
    conn.Execute sqlDelete

    ' Insertar nuevo registro
    sqlInsert = "INSERT INTO sfdc_activar_det_sw_ecbcont_peru (" & _
        "Id, Tax_Number, Opportunity_ID, Product_Code, Product_Name, " & _
        "FechaInicio, FechaFin, Fecha_Activacion, Fecha_Desactivacion, " & _
        "Codigo1, Codigo2, Codigo_Desact, Host, Sist_op_desc, " & _
        "Sist_op_version, Suscribe, Estado_Id) " & _
        "VALUES (" & _
        activacion.Id & ", '" & Replace(activacion.Tax_Number, "'", "''") & "', '" & Replace(activacion.Opportunity_ID, "'", "''") & "', '" & Replace(activacion.Product_Code, "'", "''") & "', '" & Replace(activacion.Product_Name, "'", "''") & "', " & _
        "'" & Format(activacion.FechaInicio, "yyyy-mm-dd hh:nn:ss") & "', " & _
        "'" & Format(activacion.FechaFin, "yyyy-mm-dd hh:nn:ss") & "', " & _
        "'" & Format(activacion.Fecha_Activacion, "yyyy-mm-dd hh:nn:ss") & "', " & _
        "'" & Format(activacion.Fecha_Desactivacion, "yyyy-mm-dd hh:nn:ss") & "', " & _
        "'" & Replace(activacion.Codigo1, "'", "''") & "', " & _
        "'" & Replace(activacion.Codigo2, "'", "''") & "', " & _
        "'" & Replace(activacion.Codigo_Desact, "'", "''") & "', " & _
        "'" & Replace(activacion.Host, "'", "''") & "', " & _
        "'" & Replace(activacion.Sist_op_desc, "'", "''") & "', " & _
        "'" & Replace(activacion.Sist_op_version, "'", "''") & "', " & _
        IIf(activacion.Suscribe, 1, 0) & ", " & activacion.Estado_Id & _
        ")"

    conn.Execute sqlInsert
    conn.Close

    MsgBox "Registro reemplazado exitosamente.", vbInformation
    Exit Sub

Errores:
    MsgBox "Error al reemplazar registro: " & Err.Description, vbCritical
    If Not conn Is Nothing Then If conn.State = adStateOpen Then conn.Close
End Sub



