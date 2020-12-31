Attribute VB_Name = "Globais"
Global sMaquina As String * 30
Global nLoteErroCMC7() As Integer
Global nLoteErroCB() As Integer

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character

Public Function PegarOpcaoINI(ByVal pvsSecao As String, pvsItem As String, ByVal pvsDefault As String) As String
    Dim iRet As Long
    Dim sDado As String
    Dim sDadoAux As String
    Dim i As Integer
    Dim sArquivoINI As String

    'Abrir arquivo INI conforme opção
    If pvsSecao = "Conexao" And (pvsItem = "Senha" Or pvsItem = "Usuario") Then
        sArquivoINI = App.Path & "\MDI_Conexao.INI"
    Else
        sArquivoINI = "C:\MDI_UBB\MDI_UBB.INI"
    End If

    sDado = String(255, " ")
    iRet = GetPrivateProfileString(pvsSecao, pvsItem, pvsDefault, sDado, 255, sArquivoINI)

    sDado = Trim(sDado)
    sDadoAux = ""

    For i = 1 To Len(sDado)
        If pvsSecao = "Conexao" And pvsItem = "Senha" Then
            If Asc(Mid(sDado, i, 1)) <> 0 Then
                sDadoAux = sDadoAux & Mid(sDado, i, 1)
            End If
        Else
            If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
                sDadoAux = sDadoAux & Mid(sDado, i, 1)
            End If
        End If
    Next

    If pvsSecao = "Conexao" And pvsItem = "Senha" Then
        PegarOpcaoINI = Decript(Trim(sDadoAux))
    Else
        PegarOpcaoINI = Trim(sDadoAux)
    End If

End Function

Function CriaDir(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    CriaDir = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                MkDir strPath
            End If
        End If
    Loop Until intAnchor = 0
    
    CriaDir = DirExists(strDirName)

Done:
    ChDir strOldPath

    Err = 0
End Function

Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    AddDirSep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function

Public Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub


Public Function Decript(ByVal aString As String) As String
    Dim Result(0 To 255) As Byte
    Dim Count As Long
    Dim Remainder As Integer
    Dim Quocient As Integer
    Dim Retorno As String
    Dim Key(3) As Integer
    
    ''''''''''''''''''''''''''''''''''''''
    ' Chave do algorito de criptografia  '
    ''''''''''''''''''''''''''''''''''''''
    Key(0) = 97
    Key(1) = 150
    Key(2) = 127
    Key(3) = 254
    
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

