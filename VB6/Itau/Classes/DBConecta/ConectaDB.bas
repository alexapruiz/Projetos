VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConectaDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private var_EOF As Boolean
Public Function Init(ByVal NameSTPROC As String) As Boolean

    Init = False

    'Inicializa a stored procedure.
    If SQLRPCInit%(SqlConn%, NameSTPROC, 0) = FAIL% Then Exit Function

    Init = True
End Function
Public Function ParametroIN(ByVal NomeParametro As String, ByVal ValorParametro As String, ByVal TipoParametro As Integer) As Boolean

    ParametroIN = False

    'Passagem de Parâmetro
    If FU_Parametro(NomeParametro, ValorParametro, TipoParametro) = FAIL Then Exit Function

    ParametroIN = True
End Function
Public Function ParametroOUT(ByVal NomeParametro As String, ByVal ValorParametro As String, ByVal TipoParametro As Integer) As Boolean

    ParametroOUT = False

    If FU_Parametro_Ret(NomeParametro, ValorParametro, TipoParametro) = FAIL Then Exit Function

    ParametroOUT = True
End Function
Public Function Execute() As Boolean

    Execute = False

    'Envia parametro para servidor.
    If SQLRPCSend(SqlConn) = FAIL Then Exit Function

    'Executa stored procedure.
    If SqlOk(SqlConn) = FAIL Then Exit Function

    Execute = True
End Function
Public Property Get EOF() As Variant

    EOF = var_EOF
End Property
Public Function Proximo() As Boolean

    Dim Ret As Integer

    Ret% = SqlNextRow%(SqlConn%)
    
    If Ret% = NOMOREROWS Or Ret% = FAIL Then
        var_EOF = True
    End If
End Function
Public Function LeResultados()

    Dim Ret As Integer

    Ret = SqlResults(SqlConn%)
End Function
Public Function PreencheGrid(ByRef Grid As Grid, ByVal QtdeColunas As Integer) As Boolean

    Dim i               As Integer
    Dim Situacao        As String
    Dim Ret             As Integer
    Dim x               As Integer
    Dim strAux          As String
    Dim Db              As New ConectaDB

    'Deixa grid com apenas uma linha
    Grid.Rows = 1
    Flag_Grid = True

    Db.LeResultados
    Db.Proximo

    Do Until Db.EOF
        For x = 1 To QtdeColunas
            strAux = strAux & SqlData$(SqlConn%, x) & Chr(9)
        Next x

        Grid.AddItem strAux
        
        Db.Proximo
    Loop

    'Verifica se grid possui mais de 5 linhas
    If Grid.Rows > 6 Then
        Grid.Width = 7965
    Else
        Grid.Width = 7725
    End If

    'Remove linha em branco
    If Grid.Rows > 1 Then
        Grid.RemoveItem 0
    End If
End Function
