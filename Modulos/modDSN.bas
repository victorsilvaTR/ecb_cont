Attribute VB_Name = "modDSN"
Option Explicit
'*****************************************************************************************************************'
'Print GetDatoString(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir")
'D:\ECB\ContabilidadECB
'Call SetDatoString(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir", app.Path )
'*****************************************************************************************************************'

'Contantes para crear claves
Public Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4

'Constantes para la Raiz
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

'Otras constantes
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0

'Abrir una clave
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Cerrar una clave
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Crear una clave
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'Borrar una clave
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'Recuperar un valor, sea long o string
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Crea o modifica un valor string
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
'Crea o modifica un valor long
Public Declare Function RegSetValueExL Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Borra un valor, sea long o string
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'Ajuste para eliminar una clave
Public Sub EliminarClave(Raiz As Long, Clave As String)
'El valor de retorno de la función
Dim Retorno As Long
Retorno = RegDeleteKey(Raiz, Clave)
End Sub

'Ajuste para eliminar un valor
Public Sub EliminarValor(Raiz As Long, Clave As String, Valor As String)
'El valor de retorno de la función
Dim Retorno As Long
'Manejador de la clave
Dim Manejador As Long
'Abrimos la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Eliminamos el valor buscado
Retorno = RegDeleteValue(Manejador, Valor)
'Cerramos la clave
RegCloseKey (Manejador)
End Sub

'Ajuste para recuperar una cadena
Function GetDatoString(Raiz As Long, Clave As String, Valor As String) As String
On Error Resume Next
'Manejador de la clave
Dim Manejador As Long
'Longitud de la cadena
Dim Longitud As Long
'Variable para colocar el dato
Dim Dato As String
'El valor de retorno de la función
Dim Retorno As Long
'Abrimos la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Longitud del dato
Retorno = RegQueryValueEx(Manejador, Valor, 0, 0, 0, Longitud)
'Buffer para contener la cadena y que se amolde a la longitud
Dato = String(Longitud, 0)
'Obtener el dato
Retorno = RegQueryValueEx(Manejador, Valor, 0&, REG_SZ, ByVal Dato, Longitud)
'Cerramos la clave
Retorno = RegCloseKey(Manejador)
'Quitamos el último caracter de la cadena que es el caracter de fin de linea
'Devolvemos el dato
GetDatoString = Left(Dato, Longitud - 1)
End Function

'Ajuste para recuperar un long
Function GetDatoLong(Raiz As Long, Clave As String, Valor As String) As Long
On Error Resume Next
'Manejador de la clave
Dim Manejador As Long
'Variable para contener el dato
Dim Dato As Long
'El valor de retorno de la función
Dim Retorno As Long
'Abrimos la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Obtenemos el dato
Retorno = RegQueryValueEx(Manejador, Valor, 0&, REG_DWORD, Dato, 4)
'Cerramos la clave
Retorno = RegCloseKey(Manejador)
'Devolvemos el valor númerico consultado
GetDatoLong = Dato
End Function

'Ajuste para crear una clave
Public Sub CrearClave(Raiz As Long, Clave As String)
'Manejador de la clave
Dim Manejador As Long
'No se necesita esta clave salvo para pasarla cómo parametro
Dim SA As SECURITY_ATTRIBUTES
'El valor de retorno de la función
Dim Retorno As Long
'Creamos la clave "Clave"
Retorno = RegCreateKeyEx(Raiz, Clave, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, Manejador, 0)
'Cerramos la clave que creamos
RegCloseKey (Manejador)
End Sub

'Ajuste para crear o modificar una cadena
Public Sub SetDatoString(Raiz As Long, Clave As String, Valor As String, Dato As String)
'Manejador de la clave
Dim Manejador As Long
'El valor de retorno de la función
Dim Retorno As Long
'Abrimos la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Establecemos el valor de la cadena a crear, con esto si no existe la creamos automaticamente
Retorno = RegSetValueEx(Manejador, Valor, 0, REG_SZ, Dato, Len(Dato))
'Cerramos la clave
RegCloseKey (Manejador)
End Sub

'Ajuste para crear o modificar un long
Public Sub SetDatoLong(Raiz As Long, Clave As String, Valor As String, Dato As Long)
'Manejadorde la clave
Dim Manejador As Long
'El valor de retorno de la función
Dim Retorno As Long
'Abrir la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Establecemos el valor de la cadena
Retorno = RegSetValueExL(Manejador, Valor, 0, REG_DWORD, Dato, Len(Dato))
'Cerramos la clave
RegCloseKey (Manejador)
End Sub

'Ajuste para recuperar el tip de dato de un valor
Public Function TipoValor(Raiz As Long, Clave As String, Valor As String) As Long
'Almacenara el tipo de valor que estamos consultando
Dim Tipo As Long
'Almacena la longitud del dato
Dim Longitud As Long
'Manejador de la clave
Dim Manejador As Long
'El valor de retorno de la función
Dim Retorno As Long
'Abrimos la clave
Retorno = RegOpenKeyEx(Raiz, Clave, 0, KEY_ALL_ACCESS, Manejador)
'Consultamos el tipo de dato que contiene el valor
Retorno = RegQueryValueEx(Manejador, Valor, 0, Tipo, 0, Longitud)
'Cerramos la clave
Retorno = RegCloseKey(Manejador)
'Regresamos el tipo del valor
TipoValor = Tipo
End Function

Public Sub CrearArchivoDsn(cRuta As String, servidor As String, usuario As String, password As String, bdatos As String, Autenticacion As String)
  
    Dim File As New Scripting.FileSystemObject
    
    If File.FileExists(cRuta & "\" & gsDSN & ".dsn") Then
        File.DeleteFile cRuta & "\" & gsDSN & ".dsn", True
    End If
    
    Open cRuta & "\" & gsDSN & ".dsn" For Output As #1

    Print #1, "[ODBC]"
    Print #1, "DRIVER=SQL Server"
    
    If UCase(Autenticacion) = "True" Then
        Print #1, "Trusted_Connection=YES"
    Else
        Print #1, "UID=" & usuario
    End If
    
    Print #1, "DATABASE=" & bdatos
    Print #1, "APP=Microsoft Open Database Connectivity"
    Print #1, "SERVER=" & servidor
    Print #1, "DESCRIPTION=" & cRuta
    Close #1
    
    Set File = Nothing
    
End Sub

Public Sub CrearDsn(servidor As String, usuario As String, password As String, bdatos As String, Autenticacion As String)


    Dim File As New Scripting.FileSystemObject
    Dim cRuta As String
    'windows xp
    If gsVersionWindowsMayor = 5 And gsVersionWindowsMenor = 1 Then

        cRuta = "C:\Archivos de programa\Archivos comunes\ODBC\Data Sources"
        
        If File.FolderExists(cRuta) Then
        
            Call CrearArchivoDsn(cRuta, servidor, usuario, password, bdatos, Autenticacion)
            
            Else
                cRuta = "C:\Program Files\Common Files\ODBC\Data Sources"
                
                If File.FolderExists(cRuta) Then
                
                Call CrearArchivoDsn(cRuta, servidor, usuario, password, bdatos, Autenticacion)
                
                Else
                Call CrearArchivoDsn(App.Path, servidor, usuario, password, bdatos, Autenticacion)
                Call SetDatoString(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir", App.Path)
                
            End If
        End If
    'windowsvista or windows 7
    ElseIf (gsVersionWindowsMayor = 6 And gsVersionWindowsMenor = 0) Or _
    (gsVersionWindowsMayor = 6 And gsVersionWindowsMenor = 1) Then
        Call CrearArchivoDsn(App.Path, servidor, usuario, password, bdatos, Autenticacion)
        
        Call SetDatoString(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\ODBC File DSN", "DefaultDSNDir", App.Path)
        Call SetDatoString(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & gsDSN, "Database", bdatos)
        Call SetDatoString(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & gsDSN, "LastUser", usuario)
        Call SetDatoString(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & gsDSN, "Server", servidor)
    End If
    Set File = Nothing

End Sub
