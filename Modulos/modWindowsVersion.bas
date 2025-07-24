Attribute VB_Name = "modWindowsVersion"
'------------------------------------------------------------------------------
' Mostrar información de la versión de Windows                      (11/Ago/98)
'
' Revisiones:
'   13/dic/2001 con la declaración de OSVERSIONINFOEX
'   23/May/2007 Valores para Windows Vista y uso de GetVersion
'
' ©Guillermo 'guille' Som, 1998-2001, 2007
'
' Nota:
' El ejemplo de uso de GetVersionEx seguramente está sacado de la documentación,
' pero no recuerdo exactamente de dónde y seguramente era de código de C/C++
' que convertí... pero como hace tanto tiempo, ni me acuerdo.
'
' dwMajorVersion
'   Identifies the major version number of the operating system as follows.
'   Operating System    Value
'   ----------------    -----
'   Windows 95          4
'   Windows 98          4
'   Windows Me          4
'   Windows NT 3.51     3
'   Windows NT 4.0      4
'   Windows 2000        5
'   Windows XP          5
'   Whistler            5
'   Vista/Longhorn      6
'
' dwMinorVersion
'   Identifies the minor version number of the operating system as follows.
'   Operating System    Value
'   ----------------    -----
'   Windows 95          0
'   Windows 98          10
'   Windows Me          90
'   Windows NT 3.51     51
'   Windows NT 4.0      0
'   Windows 2000        0
'   Windows XP          1
'   Whistler            1
'   Windows 2003        2 (dwMajorVersion = 5)
'   Vista/Longhorn      0 (dwMajorVersion = 6)
'
' dwBuildNumber
'   Identifies the build number of the operating system.
'
' dwPlatformId
'   Identifies the operating system platform. This member can be VER_PLATFORM_WIN32_NT.
'
' szCSDVersion
'   Contains a null-terminated string, such as "Service Pack 3", that indicates the latest Service Pack installed on the system. If no Service Pack has been installed, the string is empty.
'
' wServicePackMajor
'   Identifies the major version number of the latest Service Pack installed on the system.
'   For example, for Service Pack 3, the major version number is 3.
'   If no Service Pack has been installed, the value is zero.
'
' wServicePackMinor
'   Identifies the minor version number of the latest Service Pack installed on the system.
'   For example, for Service Pack 3, the minor version number is 0.
'
' wSuiteMask
'   A set of bit flags that identify the product suites available on the system.
'   This member can be a combination of the following values.
'   Value                               Meaning
'   -----                               -------
'   VER_SUITE_BACKOFFICE                Microsoft BackOffice components are installed.
'   VER_SUITE_DATACENTER                Windows 2000 DataCenter Server is installed.
'   VER_SUITE_ENTERPRISE                Windows 2000 Advanced Server is installed.
'   VER_SUITE_SMALLBUSINESS             Microsoft Small Business Server is installed.
'   VER_SUITE_SMALLBUSINESS_RESTRICTED  Microsoft Small Business Server is installed with the restrictive client license in force.
'   VER_SUITE_TERMINAL                  Terminal Services is installed.
'   VER_SUITE_PERSONAL                  Whistler: Whistler Personal is installed.
'
' wProductType
'   Indicates additional information about the system.
'   This member can be one of the following values.
'   Value                               Meaning
'   -----                               -------
'   VER_NT_WORKSTATION                  Windows 2000 Professional
'   VER_NT_DOMAIN_CONTROLLER            Windows 2000 domain controller
'   VER_NT_SERVER                       Windows 2000 Server
'
'            else if ( osvi.wProductType == VER_NT_SERVER )
'            {
'               if( osvi.wSuiteMask & VER_SUITE_DATACENTER )
'                  printf ( "DataCenter Server " );
'               else if( osvi.wSuiteMask & VER_SUITE_ENTERPRISE )
'                  printf ( "Advanced Server " );
'               else
'                  printf ( "Server " );
'            }
'
' wReserved
'   Reserved for future use.
'------------------------------------------------------------------------------
Option Explicit
'
'
' Declaradas en WinNT.h
Private Enum eSuiteMask
'    VER_SERVER_NT = &H80000000
'    VER_WORKSTATION_NT = &H40000000
    VER_SUITE_SMALLBUSINESS = &H1 ' Microsoft Small Business Server was once installed on the system, but may have been upgraded to another version of Windows
    VER_SUITE_ENTERPRISE = &H2 ' Windows Server 2003, Enterprise Edition, Windows 2000 Advanced Server, or Windows NT 4.0 Enterprise Edition
    VER_SUITE_BACKOFFICE = &H4 ' Microsoft BackOffice
    VER_SUITE_COMMUNICATIONS = &H8
    VER_SUITE_TERMINAL = &H10 ' Terminal Services is installed
    VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20 ' Microsoft Small Business Server is installed with the restrictive client license in force
    VER_SUITE_EMBEDDEDNT = &H40 ' Windows XP Embedded
    VER_SUITE_DATACENTER = &H80 ' Windows Server 2003, Datacenter Edition or Windows 2000 Datacenter Server
    VER_SUITE_SINGLEUSERTS = &H100 ' Terminal Services is installed, but only one interactive session is supported
    VER_SUITE_PERSONAL = &H200 ' Windows XP Home Edition
    VER_SUITE_BLADE = &H400 ' Windows Server 2003, Web Edition
    VER_SUITE_STORAGE_SERVER = &H2000 ' Windows Storage Server 2003 R2 or Windows Storage Server 2003
    VER_SUITE_COMPUTE_SERVER = &H4000 ' Windows Server 2003, Compute Cluster Edition
End Enum
'
'//
'// RtlVerifyVersionInfo() type mask bits
'//
'
Private Const VER_MINORVERSION = &H1
Private Const VER_MAJORVERSION = &H2
Private Const VER_BUILDNUMBER = &H4
Private Const VER_PLATFORMID = &H8
Private Const VER_SERVICEPACKMINOR = &H10
Private Const VER_SERVICEPACKMAJOR = &H20
Private Const VER_SUITENAME = &H40
Private Const VER_PRODUCT_TYPE = &H80
'
'//
'// RtlVerifyVersionInfo() os product type values
'//
'
Private Enum eProductType
    VER_NT_WORKSTATION = &H1
    VER_NT_DOMAIN_CONTROLLER = &H2
    VER_NT_SERVER = &H3
End Enum
'
'//
'// dwPlatformId defines:
'//
'
Private Enum ePlatformId
    VER_PLATFORM_WIN32s = 0&
    VER_PLATFORM_WIN32_WINDOWS = 1&
    VER_PLATFORM_WIN32_NT = 2&
End Enum
'
Private Type OSVERSIONINFOEX_Enum
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As ePlatformId
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As eSuiteMask
    wProductType As eProductType
    wReserved As Byte
End Type
'
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
'
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'
Private Declare Function GetVersionEx2 Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Public Sub SaberVersionWindows()
    Dim OSInfo As OSVERSIONINFO
    Dim OSInfoEx As OSVERSIONINFOEX
    Dim osvi As OSVERSIONINFOEX
    Dim ret As Long
    Dim s As String
    '
'    Move (Screen.Width - Width) \ 4, 0
    '
    ' Usando GetVersion
    ret = GetVersion
    
    Dim vMajor As Long, vMinor As Long, vBuild As Long
    
    vMajor = LoByte(LoWord(ret))
    vMinor = HiByte(LoWord(ret))

    ' A partir de NT 4 tiene el Build (no en Me/9x)
'    If ret < &H80000000 Then
'        vBuild = HiWord(ret)
'    End If
    vBuild = HiWord(ret)
    
'    Me.LabelGetVersion.Caption = "Versión: " & _
'                vMajor & "." & vMinor & "." & vBuild
'
    
    '
    OSInfo.szCSDVersion = Space$(128)
    OSInfo.dwOSVersionInfoSize = Len(OSInfo) '148
    ret = GetVersionEx(OSInfo)
    s = "MajorVersion     " & OSInfo.dwMajorVersion & vbCrLf & _
        "MinorVersion     " & OSInfo.dwMinorVersion & vbCrLf & _
        "BuildNumber      " & OSInfo.dwBuildNumber & vbCrLf & _
        "PlatformId       " & OSInfo.dwPlatformId & vbCrLf & _
        "CSDVersion       " & szTrim(OSInfo.szCSDVersion) & vbCrLf & _
        "ret              " & ret
'    Me.txtOSVersion.Text = s
    '
    '
    ' Usando OSVERSIONINFOEX
    OSInfoEx.szCSDVersion = Space$(128)
    OSInfoEx.dwOSVersionInfoSize = Len(OSInfoEx)
    ret = GetVersionEx2(OSInfoEx)
    s = "MajorVersion     " & OSInfoEx.dwMajorVersion & vbCrLf & _
        "MinorVersion     " & OSInfoEx.dwMinorVersion & vbCrLf & _
        "BuildNumber      " & OSInfoEx.dwBuildNumber & vbCrLf & _
        "PlatformId       " & OSInfoEx.dwPlatformId & vbCrLf & _
        "CSDVersion       " & szTrim(OSInfoEx.szCSDVersion) & vbCrLf & _
        "ServicePackMajor " & OSInfoEx.wServicePackMajor & vbCrLf & _
        "ServicePackMinor " & OSInfoEx.wServicePackMinor & vbCrLf & _
        "SuiteMask        " & OSInfoEx.wSuiteMask & vbCrLf & _
        "ProductType      " & OSInfoEx.wProductType & vbCrLf
        '"ret              " & ret
    s = s & vbCrLf
    '
    '
    LSet osvi = OSInfoEx
    '
    Select Case (osvi.dwPlatformId)
        Case VER_PLATFORM_WIN32_NT
            '// Test for the product.
            If (osvi.dwMajorVersion <= 4) Then
                s = s & ("Microsoft Windows NT ")
            End If
            '
            If (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 0) Then
                s = s & ("Microsoft Windows 2000 ")
            End If
            '
            If (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 1) Then
                s = s & ("Microsoft Windows XP ")
            End If
            '
            If (osvi.dwMajorVersion = 6 And osvi.dwMinorVersion = 0) Then
                s = s & ("Microsoft Windows Vista ")
            End If
            '
            If (osvi.dwMajorVersion = 6 And osvi.dwMinorVersion = 1) Then
                s = s & ("Microsoft Windows 7 ")
            End If
            '
            '// Test for product type.
            If (ret) Then
                If (osvi.wProductType = VER_NT_WORKSTATION) Then
                    If (osvi.wSuiteMask And VER_SUITE_PERSONAL) Then
                        s = s & ("Personal ")
                    Else
                        s = s & ("Professional ")
                    End If
                ElseIf (osvi.wProductType = VER_NT_SERVER) Then
                    If (osvi.wSuiteMask And VER_SUITE_DATACENTER) Then
                        s = s & ("DataCenter Server ")
                    ElseIf (osvi.wSuiteMask And VER_SUITE_ENTERPRISE) Then
                        s = s & ("Advanced Server ")
                    Else
                        s = s & ("Server ")
                     End If
                End If
            End If
            '
            '// Display version, service pack (if any), and build number.
            If (osvi.dwMajorVersion >= 4) Then
                s = s & "version " & CStr(osvi.dwMajorVersion) & "." & _
                    CStr(osvi.dwMinorVersion) & " " & _
                    (szTrim(osvi.szCSDVersion)) & _
                    " Build " & CStr(osvi.dwBuildNumber And &HFFFF)
            Else
                s = s & "version " & (szTrim(osvi.szCSDVersion)) & _
                    " Build " & CStr(osvi.dwBuildNumber And &HFFFF)
            End If
            
        Case VER_PLATFORM_WIN32_WINDOWS
            If (osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 0) Then
                s = s & "Microsoft Windows 95 "
                If (Mid$(szTrim(osvi.szCSDVersion), 2) = "C") Then _
                    s = s & ("OSR2 ")
            End If
            '
            If (osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 10) Then
                s = s & ("Microsoft Windows 98 ")
                If Mid$(szTrim(osvi.szCSDVersion), 2) = "A" Then _
                    s = s & ("SE ")
            End If
            If (osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 90) Then _
                s = s & ("Microsoft Windows Me ")
                
        Case VER_PLATFORM_WIN32s
            s = s & ("Microsoft Win32s ")
    End Select
    gsVersionWindowsMayor = osvi.dwMajorVersion
    gsVersionWindowsMenor = osvi.dwMinorVersion
    '
    '
'    Me.txtOSVersionEx.Text = s
    '
End Sub

Private Function szTrim(ByVal s As String) As String
    ' Quita los caracteres en blanco y los Chr$(0)                  (13/Dic/01)
    Dim i As Long
    '
    i = InStr(s, vbNullChar)
    If i Then
        s = Left$(s, i - 1)
    End If
    s = Trim$(s)
    
    szTrim = s
End Function

Private Function LoWord(ByVal Numero As Long) As Long
    ' Devuelve el LoWord del número pasado como parámetro
    LoWord = Numero And &HFFFF&
End Function

Private Function HiWord(ByVal Numero As Long) As Long
    ' Devuelve el HiWord del número pasado como parámetro
    HiWord = Numero \ &H10000 And &HFFFF&
End Function

Private Function LoByte(ByVal Numero As Integer) As Integer
    ' Devuelve el LoByte del número pasado como parámetro
    LoByte = Numero And &HFF
End Function

Private Function HiByte(ByVal Numero As Integer) As Integer
    ' Devuelve el HiByte del número pasado como parámetro
    HiByte = Numero \ &H100 And &HFF
End Function




