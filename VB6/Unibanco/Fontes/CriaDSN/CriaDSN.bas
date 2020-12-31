Attribute VB_Name = "Module1"
Option Explicit

'Função de Leitura de .INI
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Variáveis de Trabalho
Private Server As String
Private DataBase As String
Private DataBaseBackup As String
Private NetWork As String
Private Name As String
Private NameBackup As String

Sub CriaDSN()

  Dim strAttribs As String

  strAttribs = "OemToAnsi=No" _
  & Chr$(13) & "SERVER=" & Server _
  & Chr$(13) & "Network=" & NetWork _
  & Chr$(13) & "Database=" & DataBase _
  & Chr$(13) & "USEPROCFORPREPARE=No"

  On Error GoTo ErroRegistro

  rdoErrors.Clear

  'Cria DSN Atual
  rdoEngine.rdoRegisterDataSource Name, _
         "SQL Server", True, strAttribs

  MsgBox "Data Source '" & Name & "' criado com sucesso.", vbInformation + vbOKOnly, App.Title

  On Error GoTo 0

  If Len(Trim(DataBaseBackup)) <> 0 And Len(Trim(NameBackup)) <> 0 Then
    'Cria DSN Backup
    strAttribs = "OemToAnsi=No" _
    & Chr$(13) & "SERVER=" & Server _
    & Chr$(13) & "Network=" & NetWork _
    & Chr$(13) & "Database=" & DataBaseBackup _
    & Chr$(13) & "USEPROCFORPREPARE=No"

    On Error GoTo ErroRegistro

    rdoErrors.Clear

    rdoEngine.rdoRegisterDataSource NameBackup, _
           "SQL Server", True, strAttribs

    MsgBox "Data Source '" & NameBackup & "' criado com sucesso.", vbInformation + vbOKOnly, App.Title

  End If

  Exit Sub

ErroRegistro:
  MsgBox rdoErrors(0).Description
End Sub
Sub LeINI()

  Dim iRet As Long
  Dim sDado As String
  Dim sDadoAux As String
  Dim i As Integer
  
  'Name Atual
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "NAME", "MDI_Ubb", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  Name = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          Name = Name & Mid(sDado, i, 1)
      End If
  Next

  'Name Backup
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "NAME_BACKUP", "MDI_Ubb", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  NameBackup = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          NameBackup = NameBackup & Mid(sDado, i, 1)
      End If
  Next

  'Server
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "SERVER", "server_nt", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  Server = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          Server = Server & Mid(sDado, i, 1)
      End If
  Next

  'Database Atual
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "DATABASE", "MDI_Ubb", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  DataBase = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          DataBase = DataBase & Mid(sDado, i, 1)
      End If
  Next

  'Database Backup
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "DATABASE_BACKUP", "MDI_Ubb", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  DataBaseBackup = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          DataBaseBackup = DataBaseBackup & Mid(sDado, i, 1)
      End If
  Next

  'Network
  sDado = String(255, " ")
  iRet = GetPrivateProfileString("Conexao", "NETWORK", "dbmssocn", sDado, 255, App.Path & "\CriaDSN.INI")

  sDado = Trim(sDado)
  NetWork = ""

  For i = 1 To Len(sDado)
      If Asc(Mid(sDado, i, 1)) >= 32 And Asc(Mid(sDado, i, 1)) <= 122 Then
          NetWork = NetWork & Mid(sDado, i, 1)
      End If
  Next
End Sub
Sub Main()

  Call LeINI
  
  Call CriaDSN
End Sub
