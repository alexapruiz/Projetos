Attribute VB_Name = "basDBMDI"
Option Explicit

Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
Function formataValor(ByVal pValor As String, Optional ByVal DivideCem As Boolean) As String
   
    Dim sValor As String
    
    If DivideCem Then
        sValor = Format(pValor / 100, "000000000000.00")
    Else
        sValor = Format(pValor, "000000000000.00")
    End If
    
    Mid(sValor, InStr(sValor, ","), 1) = "."
       
    formataValor = sValor
   
End Function
Function GetComputerName() As String
    ' Set or retrieve the name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = Left$(strBuffer, lngLen)
    Else
        GetComputerName = ""
    End If
End Function
Function GetSource(ByVal pstrSource As String, ByVal pstrProc As String, ByVal pstrMod As String) As String
    Dim strFront As String
    Dim strBack As String
  
    strBack = pstrProc & "@" & GetComputerName()
  
    If Left$(pstrMod, (InStr(1, pstrMod, ":") - 1)) = pstrSource Then
        strFront = pstrSource & Mid$(pstrMod, InStr(1, pstrMod, ":"))
    Else
        strFront = "|" & pstrMod
    End If
  
    GetSource = strFront & strBack
End Function
Function SetErrSource(modName As String, procName As String) As String

    SetErrSource = modName & "." & procName & " Estação: " & GetComputerName()

End Function




