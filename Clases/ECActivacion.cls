VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECActivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim ActivationDt As ECCOMActivacion.ActivationData  'Clase Objeto que representa los datos de una Activacion en la Base datos Local!
Public ManagerDt As New ECCOMActivacion.ManagerData  'Clase que contiene los metodos de acceso a Base datos Local!


Public Function ExistDataBase() As Boolean
 On Error GoTo Errores
 ExistDataBase = ManagerDt.ExistDataBase()
   
 Exit Function
Errores:
    ExistDataBase = False
    MsgBox Err.Description, vbCritical
End Function

Public Function GetActivacion() As Variant
On Error GoTo Errores
 Set GetActivacion = ManagerDt.GetActivationData()
Exit Function
Errores:
    GetActivacion = Null
    MsgBox "Ocurrió un error: " & Err.Number & " - " & Err.Description, vbExclamation
End Function


Public Function LoadActivacion(act As ActivationData) As Boolean
On Error GoTo Errores
 LoadActivacion = ManagerDt.LoadActivacion(act)

Exit Function
Errores:
    LoadActivacion = False
    MsgBox Err.Description, vbCritical
End Function

Public Function GetPathInstallation()
On Error GoTo Errores
GetPathInstallation = ManagerDt.GetPathInstallation()

Exit Function
Errores:
    GetPathInstallation = Null
    MsgBox Err.Description, vbCritical
End Function


Public Function IsEmptyTable(tableName As String) As Boolean
On Error GoTo Errores
IsEmptyTable = ManagerDt.IsEmptyTable(tableName)

Exit Function
Errores:
    IsEmptyTable = False
    MsgBox Err.Description, vbCritical
End Function

Public Function OpenAppLicencias() As Boolean
On Error GoTo Errores
Dim rutaAppLicencia As String
    rutaAppLicencia = ManagerDt.GetPathInstallation() & "\LicenciasApp\EC.Activacion.exe"
    
    Shell Chr(34) & rutaAppLicencia & Chr(34), vbNormalFocus
    OpenAppLicencias = True
    Exit Function
Errores:
    OpenAppLicencias = False
    MsgBox Err.Description, vbCritical
End Function

Public Function OpenAppUpdates() As Boolean
On Error GoTo Errores
Dim rutaAppUpdates As String
    rutaAppUpdates = ManagerDt.GetPathInstallation() & "\UpdatesApp\ECUpdate.exe"
    
    Shell Chr(34) & rutaAppUpdates & Chr(34), vbNormalFocus
    OpenAppUpdates = True
    Exit Function
Errores:
    OpenAppUpdates = False
    MsgBox Err.Description, vbCritical
End Function


Public Function WSHasConnection() As Boolean

On Error GoTo Errores

    Dim WSHTTPCont As String
    WSHTTPCont = "http://activacion.thomsonreuters.cl/WSECActivacion/Service.asmx"

    Dim HttpRequest As Object
    Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    
'Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
    Dim payload As String
    payload = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
              "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
              "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
              "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
              "<soap:Body>" & _
                "<Test xmlns=""http://tempuri.org/""/>" & _
              "</soap:Body>" & _
              "</soap:Envelope>"

    With HttpRequest
        .Open "POST", WSHTTPCont, False
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://tempuri.org/Test"
        .send payload
    End With

   Dim xmldoc As Object
    Set xmldoc = CreateObject("MSXML2.DOMDocument")
    xmldoc.async = False
    xmldoc.LoadXML HttpRequest.responseText

    Dim node As Object
    Set node = xmldoc.SelectSingleNode("//TestResult")

    If Not node Is Nothing Then
    If node.Text = "OK" Then
        WSHasConnection = True
     End If
    Else
        WSHasConnection = False
    End If

    Exit Function

Errores:
    WSHasConnection = False
    MsgBox Err.Description, vbCritical
End Function

Public Function GetWSActivation(taxNumber As String, Host As String) As ECCOMActivacion.ActivationData
    On Error GoTo Errores

    ' Parámetros que vendrán desde la base de datos
    Dim tokenRequestId As String
    Dim json As Object
    tokenRequestId = "KlIybTVuMnIxYzM0bjJzV1NwMSRzdzRyZCQuMjAwMF9d"

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

'    If node Is Nothing Then
'    MsgBox "Error: No se encontró el nodo GetActivationCurrentResult."
'    ElseIf Trim(node.Text) = "" Or LCase(Trim(node.Text)) = "null" Then
'    MsgBox "El servicio respondió con un valor nulo o vacío."
'    Else


    Dim cleanJson As String
    cleanJson = node.Text
    ' Reemplazar comillas tipográficas por comillas normales
    cleanJson = Replace(cleanJson, "“", """")
    cleanJson = Replace(cleanJson, "”", """")

 
    ' También quitar espacios innecesarios al inicio/final
    cleanJson = Trim(cleanJson)
    
    Set json = JsonConverter.ParseJson(cleanJson)
    Set GetWSActivation = ObtenerObjetoDesdeJson(json)
       
    
    'GetDataService = cleanJson

'    End If

Exit Function

Errores:
    Exit Function
    'GetDataService = Nothing
End Function

Public Function ActualizacionDisponible(taxNumber As String, versionApp As String, nSubida As Integer) As Boolean

On Error GoTo Errores
    Dim cleanedVersion As String
    Dim headerValue() As String
    Dim WSHTTPCont As String
    WSHTTPCont = "http://activacion.thomsonreuters.cl/WSECActivacion/Service.asmx"
    
    'Limpia caracteres incorrectos
    cleanedVersion = Replace(versionApp, " ", "")
    cleanedVersion = Replace(versionApp, "v", "")
    cleanedVersion = Replace(versionApp, "V", "")
        
    'Crear solicitud SOAP
    Dim HttpRequest As Object
    Dim objStream
    Set HttpRequest = CreateObject("MSXML2.ServerXMLHTTP")

    Dim payload As String
    payload = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
              "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
              "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
              " xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
              "<soap:Body>" & _
                "<DowloadInstallerController xmlns=""http://tempuri.org/"">" & _
                  "<tax>" & taxNumber & "</tax>" & _
                  "<subida>" & nSubida & "</subida>" & _
                  "<versionInstall>" & cleanedVersion & "</versionInstall>" & _
                "</DowloadInstallerController>" & _
              "</soap:Body>" & _
              "</soap:Envelope>"

    'Enviar solicitud
    With HttpRequest
        .Open "POST", WSHTTPCont, False
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://tempuri.org/DowloadInstallerController"
        .send payload
    End With

    If HttpRequest.Status = 200 Then
       headerValue = Split(HttpRequest.getResponseHeader("Content-Disposition"), "=")
       
       Set objStream = CreateObject("ADODB.Stream")
       objStream.Open
       objStream.Type = 1

       objStream.Write HttpRequest.responseBody
       objStream.Position = 0

       objStream.SaveToFile GetDownloadedFolderPath & "\\" & headerValue(1), 2
       objStream.Close
       Set objStream = Nothing
       
       ActualizacionDisponible = True
     Else
       ActualizacionDisponible = False
     End If

Exit Function

Errores:
    Exit Function
    ActualizacionDisponible = False

End Function


Private Function ObtenerObjetoDesdeJson(json As Object) As ECCOMActivacion.ActivationData
    Dim act As New ECCOMActivacion.ActivationData
    
    On Error GoTo Err
    act.Id = json("ID")
    act.op = json("OP")
    act.Tax_Number = json("TAX_NUMBER")
    act.Product_Code = json("PRODUCT_CODE")
    act.Product_Name = json("PRODUCT_NAME")
    act.FECHA_INI = FechaParse(json("FECHA_INI"))
    act.FECHA_FIN = FechaParse(json("FECHA_FIN"))
    act.Fecha_Activacion = FechaParse(json("FECHA_ACTIVACION"))
    act.Fecha_Desactivacion = FechaParse(json("FECHA_DESACTIVACION"))
    act.Codigo1 = json("CODIGO1")
    act.Codigo2 = json("CODIGO2")
    act.Host = json("HOST")
    act.SISOP = json("SISOP")
    act.SISVER = json("SISVER")
    act.Suscribe = json("SUSCRIBE")
    act.Estado_Id = json("ESTADO_ID")
    
    Set ObtenerObjetoDesdeJson = act
Exit Function
Err:
    Set ObtenerObjetoDesdeJson = Nothing
End Function



Public Function FechaParse(fechaValue As Variant) As Date
On Error GoTo Err
    If Not IsNull(fechaValue) And Trim(fechaValue) <> "" Then
        ' Reemplazar la "T" por un espacio
        fechaValue = Replace(fechaValue, "T", " ")
        ' Convertir a fecha
        FechaParse = CDate(fechaValue)
    Else
        ' Manejo de fecha vacía o nula
        FechaParse = #1/1/100#
    End If
Exit Function
Err:
    FechaParse = #1/1/100#
End Function


Function RegKeyRead(i_RegKey As String) As String
    
    Dim myWS As Object

    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = myWS.RegRead(i_RegKey)
    
End Function

Public Function Replace(strExpression As Variant, strSearch As String, StrReplace As String) As String

    Dim lngStart As Long
    
    If IsNull(strExpression) Then Exit Function
    
    lngStart = 1
    While InStr(lngStart, strExpression, strSearch) <> 0
        lngStart = InStr(lngStart, strExpression, strSearch)
        strExpression = Left(strExpression, lngStart - 1) & StrReplace & Mid(strExpression, lngStart + Len(strSearch))
        lngStart = lngStart + Len(StrReplace)
    Wend

    Replace = strExpression
    
End Function

Function GetDownloadedFolderPath() As String

    GetDownloadedFolderPath = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}")
    GetDownloadedFolderPath = Replace(GetDownloadedFolderPath, "%USERPROFILE%", Environ$("USERPROFILE"))

End Function
