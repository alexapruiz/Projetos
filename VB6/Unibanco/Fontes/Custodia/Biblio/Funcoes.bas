Attribute VB_Name = "Funcoes"
Option Explicit

'''''''''''''''''''''''''''''''
'   Acerta Tipo de Dados
'''''''''''''''''''''''''''''''
Public Function AcertaTD(ByVal pStr As String) As String


    Dim sStr            As String
    Dim sStrRetorno     As String
    Dim iPos            As Integer

    iPos = 1
    Do While iPos <> 0
    
        iPos = InStr(pStr, ",")
        
        sStr = Mid(pStr, 1, IIf(iPos = 0, Len(pStr), iPos - 1))
        pStr = Mid(pStr, iPos + 1)
        
        If Not IsNumeric(sStr) Then
            sStr = getParams(sStr)
        End If
        
        If Len(sStrRetorno) > 0 Then sStrRetorno = sStrRetorno & ","
        
        sStrRetorno = sStrRetorno & sStr
    Loop

    AcertaTD = sStrRetorno

End Function

Public Function tiraAOA(ByVal prmStr As String) As String
    
    Dim l_1 As String
    Dim l_2 As String
    

    Do While InStr(prmStr, "'")
        l_1 = Mid(prmStr, 1, InStr(prmStr, "'") - 1)
        l_2 = Mid(prmStr, InStr(prmStr, "'") + 1)
        prmStr = l_1 & l_2
    Loop
    Do While InStr(prmStr, """")
        l_1 = Mid(prmStr, 1, InStr(prmStr, """") - 1)
        l_2 = Mid(prmStr, InStr(prmStr, """") + 1)
        prmStr = l_1 & l_2
    Loop

    tiraAOA = prmStr
    
End Function


Public Function getParams(ParamArray prmArgs()) As String

    Dim l_iI    As Integer
    Dim l_sStr  As String
    
    For l_iI = 0 To UBound(prmArgs)
        If Len(l_sStr) > 0 Then l_sStr = l_sStr & ","
        
        If IsMissing(prmArgs(l_iI)) Then
            l_sStr = l_sStr
        Else
            l_sStr = l_sStr & atoa(prmArgs(l_iI))
        End If
    Next l_iI

    getParams = l_sStr

End Function

Public Function Aspas(ByVal pStr As String) As String


    pStr = """" & pStr & """"
    
    Aspas = pStr

End Function

Public Function atoa(ByVal prmStrToSearch_ As String) As String
    
    
    Dim lclCI           As String
    Dim lclNomeLayout   As String
    
    lclNomeLayout = prmStrToSearch_
    
    If InStr(lclNomeLayout, "'") Then
        lclCI = """"
    Else
        lclCI = "'"
    End If

    
    atoa = lclCI & prmStrToSearch_ & lclCI
    
End Function


