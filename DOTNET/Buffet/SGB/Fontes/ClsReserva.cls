VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Inserir(ByVal DATA_RES As String, ByVal HR_RES As String, ByVal NOM_CLI As String, ByVal TEL1_CLI As String, ByVal TEL2_CLI As String, ByVal DATA_FIM_RES As String, ByVal Db As ADODB.Connection) As Integer

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    On Error GoTo Inserir_Reserva_Erro

    sSql = "SELECT * FROM RESERVA WHERE DATA_RES = #" & DATA_RES
    sSql = sSql & "# AND HR_RES = '" & Left(HR_RES, 1) & "'"
    Rs.Open sSql, Db, adOpenDynamic, 1

    If Not Rs.EOF Then
        'Já existe Reserva
        MsgBox "Já existe uma reserva para este dia / hora"
    Else
        Set Rs = Nothing

        sSql = "INSERT INTO RESERVA (DATA_RES,HR_RES,NOM_CLI,TEL1_CLI,TEL2_CLI,DATA_FIM_RES) VALUES ("
        sSql = sSql & "#" & Format(DATA_RES, "mm/dd/yyyy") & "#,'"
        sSql = sSql & Left(HR_RES, 1) & "','"
        sSql = sSql & NOM_CLI & "','"
        sSql = sSql & TEL1_CLI & "','"
        sSql = sSql & TEL2_CLI & "','"
        sSql = sSql & Format(DATA_FIM_RES, "mm/dd/yyyy") & "')"
    
        Rs.Open sSql, Db, adOpenDynamic, 1

        Inserir = True
    End If

    Exit Function

Inserir_Reserva_Erro:
    MsgBox "Erro ao inserir os dados da reserva", vbExclamation, "SGB"
    Inserir = False
End Function
Public Function Atualizar(ByVal ID_RES As String, ByVal DATA_RES As String, ByVal HR_RES As String, ByVal NOM_CLI As String, ByVal TEL1_CLI As String, ByVal TEL2_CLI As String, ByVal DATA_FIM_RES As String, ByVal Db As ADODB.Connection) As Integer

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    On Error GoTo Atualizar_Reserva_Erro

    sSql = "UPDATE RESERVA SET DATA_RES = #" & DATA_RES
    sSql = sSql & "# , HR_RES = '" & HR_RES
    sSql = sSql & "' , NOM_CLI = '" & NOM_CLI
    sSql = sSql & "' , TEL1_CLI = '" & TEL1_CLI
    sSql = sSql & "' , TEL2_CLI = '" & TEL2_CLI
    sSql = sSql & "' , DATA_FIM_RES = #" & DATA_FIM_RES
    sSql = sSql & "# WHERE ID_RES = " & ID_RES

    Rs.Open sSql, Db, adOpenDynamic, 1

    Atualizar = True
    Exit Function

Atualizar_Reserva_Erro:
    MsgBox "Erro ao atualizar os dados da reserva", vbExclamation, "SGB"
    Atualizar = False
End Function
Public Function Excluir(ByVal ID_RES As String, ByVal Db As ADODB.Connection) As Integer

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    On Error GoTo Atualizar_Reserva_Erro

    sSql = "DELETE FROM RESERVA WHERE ID_RES = " & ID_RES

    Rs.Open sSql, Db, adOpenDynamic, 1

    Excluir = True
    Exit Function

Atualizar_Reserva_Erro:
    MsgBox "Erro ao Excluir a reserva", vbExclamation, "SGB"
    Excluir = False
End Function
Public Function Selecionar(ByVal DATA_RES As String, ByVal HORA_INI_FESTA As String, ByVal Db As ADODB.Connection) As Integer

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    On Error GoTo Atualizar_Reserva_Erro

    'Verificar se é almoço ou noite
    Select Case HORA_INI_FESTA
        Case Is < 7
            HORA_INI_FESTA = "A"

        Case Is > 7
            HORA_INI_FESTA = "N"
    End Select

    sSql = "SELECT * FROM RESERVA "
    sSql = sSql & " WHERE DATA_RES = #" & Format(DATA_RES, "mm/dd/yyyy") & "#"
    sSql = sSql & " AND HR_RES = '" & HORA_INI_FESTA & "'"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    If Not Rs.EOF Then
        Selecionar = True
    Else
        Selecionar = False
    End If
    Exit Function

Atualizar_Reserva_Erro:
    MsgBox "Erro ao inserir os dados da reserva", vbExclamation, "SGB"
    Selecionar = False
End Function
