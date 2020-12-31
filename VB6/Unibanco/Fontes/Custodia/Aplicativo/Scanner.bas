Attribute VB_Name = "Scanner"
'Versao do windows
 Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
 
 Public Const VER_PLATFORM_WIN32_NT = 2
 Public Const VER_PLATFORM_WIN32_WINDOWS = 1
 
 Public Enum enumWINVERSION
     eWIN_ER = 0     'Não foi possivel obter a versao do sistema operacional
     eWIN_NT = 1     'Windows NT
     eWIN_9X = 2     'Windows 95, 98
 End Enum

 Public Type OSVERSIONINFO
     dwOSVersionInfoSize         As Long
     dwMajorVersion              As Long
     dwMinorVersion              As Long
     dwBuildNumber               As Long
     dwPlatformId                As Long
     szCSDVersion                As String * 128      '  Maintenance string for PSS usage
 End Type

'Diretorio do Windows
 Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
'Diretoria de sistema
 Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
'Retornos de leitura do scanner
 Public Enum enumRetornoLeitura
     eLeituraFim = 0    'Final de cheques no alimentador
     eLeituraOK = 1     'Leitura Terminada
     eLeituraEsc = 2    'Teclou Esc
     eLeituraFalha = 3  'Falha na Leitura
     eTimeOut = 4       'Saiu por Time Out
     eErro = 9          'Erro no módulo de leitura
 End Enum
 
'Verifica no register do windows, verifica se driver LA93 esta instalado
 Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long
     
 Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef _
    lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
    
 Enum REGToolRootTypes
    HK_CLASSES_ROOT = &H80000000
    HK_CURRENT_USER = &H80000001
    HK_LOCAL_MACHINE = &H80000002
    HK_USERS = &H80000003
    HK_PERFORMANCE_DATA = &H80000004
    HK_CURRENT_CONFIG = &H80000005
    HK_DYN_DATA = &H80000006
 End Enum
    
 Const ERROR_SUCCESS = 0
 Const gsSLASH_BACKWARD As String = "\"
 Const sKEY As String = "Software\Vips France\VipsDrv\3.12"
 Const sValue As String = "Serial"    ' "Company"
    
'Reg Key Security Options...
 Const READ_CONTROL = &H20000
 Const KEY_QUERY_VALUE = &H1
 Const KEY_SET_VALUE = &H2
 Const KEY_CREATE_SUB_KEY = &H4
 Const KEY_ENUMERATE_SUB_KEYS = &H8
 Const KEY_NOTIFY = &H10
 Const KEY_CREATE_LINK = &H20
 Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
 Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
 Const KEY_EXECUTE = KEY_READ
 Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE _
                           + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS _
                           + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Function VerRegSerialLA93() As String
    Dim hKey As Long
    Dim KeyValType As Long
    Dim sTmp As String
    Dim KeyValSize As Long

    If RegOpenKeyEx(HK_LOCAL_MACHINE, sKEY, 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function

    sTmp = String$(1024, 0)
    KeyValSize = 1024

    If RegQueryValueEx(hKey, sValue, 0, KeyValType, sTmp, KeyValSize) <> ERROR_SUCCESS Then Exit Function
    
    If (Asc(Mid(sTmp, KeyValSize, 1)) = 0) Then
        sTmp = Left(sTmp, KeyValSize - 1)
    Else
        sTmp = Left(sTmp, KeyValSize)
    End If
    
    VerRegSerialLA93 = sTmp

End Function

 
Public Function GetWinVersion() As enumWINVERSION
    Dim Ret         As Long
    Dim strct       As OSVERSIONINFO
    
    GetWinVersion = eWIN_ER
    
    strct.dwOSVersionInfoSize = Len(strct)
    
    Ret = GetVersionEx(strct)
    
    If Ret <> 0 Then
        GetWinVersion = IIf(strct.dwPlatformId = VER_PLATFORM_WIN32_NT, eWIN_NT, eWIN_9X)
    End If
End Function
Public Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(255)
    
    If GetWindowsDirectory(strBuf, 255) > 0 Then
        GetWindowsDir = Trim(Replace(strBuf, Chr(0), ""))
    Else
        GetWindowsDir = vbNullString
    End If
End Function
Public Function GetWindowsSys() As String
    Dim strBuf As String

    strBuf = Space$(255)
    
    If GetSystemDirectory(strBuf, 255) > 0 Then
        GetWindowsSys = Trim(Replace(strBuf, Chr(0), ""))
    Else
        GetWindowsSys = vbNullString
    End If
End Function
Public Function AchaScannerDLL(ByVal Scanner As Integer) As Boolean

Dim achou As Boolean
Dim DirSys As String

    DirSys = GetWindowsSys()
    
    If Scanner = 1 Then
        If GetWinVersion = eWIN_9X Then
            achou = IIf(Dir(App.path + "\DTC329X.DLL") <> "", True, IIf(Dir(DirSys + "\DTC329X.DLL") <> "", True, False))
        ElseIf GetWinVersion = eWIN_NT Then
            achou = IIf(Dir(App.path + "\DTC32NT.DLL") <> "", True, IIf(Dir(DirSys + "\DTC32NT.DLL") <> "", True, False))
        End If
    ElseIf Scanner = 2 Then
        achou = IIf(Dir(App.path + "\LA93.DLL") <> "", True, IIf(Dir(DirSys + "\LA93.DLL") <> "", True, False))
    End If

AchaScannerDLL = achou

End Function

