Attribute VB_Name = "Cripto"
Option Explicit
Public Key(3) As Integer

Public Function Encript(ByVal aString As String) As String
    Dim Result(0 To 255) As Byte
    Dim Count As Long
    Dim Remainder As Integer
    Dim Quocient As Integer
    Dim Retorno As String
    
    Quocient = Len(aString) \ 4
    Remainder = Len(aString) Mod 4
    For Count = 0 To Quocient - 1
        Result(Count * 4 + 0) = (Asc(Mid(aString, Count * 4 + 1, 1)) + Key(0)) Mod 255
        Result(Count * 4 + 1) = (Asc(Mid(aString, Count * 4 + 2, 1)) + Key(1)) Mod 255
        Result(Count * 4 + 2) = (Asc(Mid(aString, Count * 4 + 3, 1)) + Key(2)) Mod 255
        Result(Count * 4 + 3) = (Asc(Mid(aString, Count * 4 + 4, 1)) + Key(3)) Mod 255
        
        Result(Count * 4 + 0) = Result(Count * 4 + 0) Xor Key(0)
        Result(Count * 4 + 1) = Result(Count * 4 + 1) Xor Key(1)
        Result(Count * 4 + 2) = Result(Count * 4 + 2) Xor Key(2)
        Result(Count * 4 + 3) = Result(Count * 4 + 3) Xor Key(3)
    Next
    If Remainder > 0 Then
        For Count = 0 To Remainder - 1
            Result(Quocient * 4 + Count) = (Asc(Mid(aString, Quocient * 4 + Count + 1, 1)) + Key(Count)) Mod 255
            Result((Quocient * 4) + Count) = Result(Quocient * 4 + Count) Xor Key(Count)
        Next
    End If
    For Count = 0 To Len(aString) - 1
        Retorno = Retorno + Chr(Result(Count))
    Next
    Encript = Retorno
    
End Function

Public Function Decript(ByVal aString As String) As String
    Dim Result(0 To 255) As Byte
    Dim Count As Long
    Dim Remainder As Integer
    Dim Quocient As Integer
    Dim Retorno As String
    
    Quocient = Len(aString) \ 4
    Remainder = Len(aString) Mod 4
    For Count = 0 To Quocient - 1
        Result(Count * 4 + 0) = Asc(Mid(aString, Count * 4 + 1, 1)) Xor Key(0)
        Result(Count * 4 + 1) = Asc(Mid(aString, Count * 4 + 2, 1)) Xor Key(1)
        Result(Count * 4 + 2) = Asc(Mid(aString, Count * 4 + 3, 1)) Xor Key(2)
        Result(Count * 4 + 3) = Asc(Mid(aString, Count * 4 + 4, 1)) Xor Key(3)
        
        Result(Count * 4 + 0) = (Result(Count * 4 + 0) - Key(0) + 255) Mod 255
        Result(Count * 4 + 1) = (Result(Count * 4 + 1) - Key(1) + 255) Mod 255
        Result(Count * 4 + 2) = (Result(Count * 4 + 2) - Key(2) + 255) Mod 255
        Result(Count * 4 + 3) = (Result(Count * 4 + 3) - Key(3) + 255) Mod 255
    Next
    If Remainder > 0 Then
        For Count = 0 To Remainder - 1
            Result(Quocient * 4 + Count) = Asc(Mid(aString, Quocient * 4 + Count + 1, 1)) Xor Key(Count)
            Result(Quocient * 4 + Count) = (Result(Quocient * 4 + Count) - Key(Count) + 255) Mod 255
        Next
    End If
    For Count = 0 To Len(aString) - 1
        Retorno = Retorno + Chr(Result(Count))
    Next
    Decript = Retorno
End Function

