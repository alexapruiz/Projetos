Attribute VB_Name = "Funções"
'Declaração DLL - WINNT
 Declare Function DTCNT_SetCommPort Lib "DTC32NT.DLL" (ByVal iCommPort As Long) As Long
 Declare Function DTCNT_SetBaudRate Lib "DTC32NT.DLL" (ByVal iBaudRate As Long) As Long
 Declare Function DTCNT_SetWordLength Lib "DTC32NT.DLL" (ByVal iWordLength As Long) As Long
 Declare Function DTCNT_SetParity Lib "DTC32NT.DLL" (ByVal iParity As Long) As Long
 Declare Function DTCNT_SetStopBits Lib "DTC32NT.DLL" (ByVal iStopBits As Long) As Long
 Declare Function DTCNT_Init Lib "DTC32NT.DLL" () As Long
 Declare Function DTCNT_DeInit Lib "DTC32NT.DLL" () As Long
 Declare Function DTCNT_Read Lib "DTC32NT.DLL" (ByVal String128 As String) As Long
 
'Declaração DLL - WIN9X
 Declare Function DTC9X_SetCommPort Lib "DTC329X.DLL" (ByVal iCommPort As Long) As Long
 Declare Function DTC9X_SetBaudRate Lib "DTC329X.DLL" (ByVal iBaudRate As Long) As Long
 Declare Function DTC9X_SetWordLength Lib "DTC329X.DLL" (ByVal iWordLength As Long) As Long
 Declare Function DTC9X_SetParity Lib "DTC329X.DLL" (ByVal iParity As Long) As Long
 Declare Function DTC9X_SetStopBits Lib "DTC329X.DLL" (ByVal iStopBits As Long) As Long
 Declare Function DTC9X_Init Lib "DTC329X.DLL" () As Long
 Declare Function DTC9X_DeInit Lib "DTC329X.DLL" () As Long
 Declare Function DTC9X_Read Lib "DTC329X.DLL" (ByVal String128 As String) As Long

 Declare Function LA93_SetCommPort Lib "LA93.DLL" (ByVal lngComPort As Long) As Long
 Declare Function LA93_SetNumEscaninhos Lib "LA93.DLL" (ByVal lngNumEscaninhos As Long) As Long
 Declare Function LA93_Init Lib "LA93.DLL" () As Long
 Declare Function LA93_DeInit Lib "LA93.DLL" () As Long
 'Declare Function LA93_Read Lib "LA93.DLL" (ByVal strCampo1 As String, ByVal strCampo2 As String, ByVal strCampo3 As String) As Long
 Declare Function LA_ClearBuffer Lib "LA93.DLL" () As Long
 Declare Function LA93_Read Lib "LA93.DLL" (ByVal strCampo1 As String) As Long
 Declare Function LA93_Eject Lib "LA93.DLL" (ByVal lngEscaninho As Long) As Long
 Declare Function LA93_GetErrorText Lib "LA93.DLL" (ByVal lngNumErro As Long, ByVal strErro As String, ByVal lngTamMax As Long) As Long

 Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

