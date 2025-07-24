Attribute VB_Name = "modLocaleInfo"
 
Public Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Public Const CurrentUser As String = "HKEY_CURRENT_USER"
  
'------------------------------------------
'    Configuraci�n de la MONEDA
'------------------------------------------
' N�mero de d�gitos decimales para la moneda: Leer_Dato(CurrentUser, "iCurrDigits")
' S�mbolo de Moneda: Leer_Dato(CurrentUser, "sCurrency")
' S�mbolo del separador decimal para la moneda: Leer_Dato(CurrentUser, "sMonDecimalSep")
' S�mbolo del separador de miles para la moneda: Leer_Dato(CurrentUser, "sMonThousandSep")
'------------------------------------------
'    Configuraci�n de los NUMEROS
'------------------------------------------
' N�mero de d�gitos decimales: Leer_Dato(CurrentUser, "iDigits")
' S�mbolo del separador de miles: Leer_Dato(CurrentUser, "sThousand")
' S�mbolo del separador decimal: Leer_Dato(CurrentUser, "sDecimal")
'------------------------------------------
'    Configuraci�n de HORA y FECHA
'------------------------------------------
' Formato de hora : Leer_Dato(CurrentUser, "sTimeFormat")
' S�mbolo separador de HORA: Leer_Dato(CurrentUser, "sTime")
' Formato de Fecha Corta: Leer_Dato(CurrentUser, "sShortDate")
' Formato de Fecha Larga: Leer_Dato(CurrentUser, "sLongDate")
  
' Funci�n que lee el valor del registro en la rama Control Panel\International
Public Function Leer_Dato(Principal As String, Valor As String) As String
'Variable para acceder al Registro mediante Wsh - Windows Scripting Host
    Dim O_Registro As New WshShell
    Leer_Dato = O_Registro.RegRead(Principal & "\Control Panel\International\" & Valor)
    Set O_Registro = Nothing
End Function
