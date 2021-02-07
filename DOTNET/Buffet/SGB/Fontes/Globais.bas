Attribute VB_Name = "Globais"
Public Db As New ADODB.Connection
Public Sub Main()

    Db.Open "SGB"

    Login.Show 1
End Sub
Public Function CarregaCombo(ByRef Combo As Object, ByVal TABELA As String, CAMPO1 As String, CAMPO2 As String, Optional WHERE As String) As Integer

    Dim x As Integer
    Dim Rs As New ADODB.Recordset

    sSql = "Select " & CAMPO1 & " as CODIGO ," & CAMPO2 & " as DESCRICAO from " & TABELA

    If Len(Trim(WHERE)) > 0 Then
        sSql = sSql & WHERE
    End If
    If Len(Trim(CAMPO2)) > 0 Then
        sSql = sSql & " ORDER BY " & CAMPO2
    End If
    Rs.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Do Until Rs.EOF
        Combo.AddItem Rs(1).Value
        Combo.ItemData(Combo.NewIndex) = Rs(0).Value
        Rs.MoveNext
    Loop
End Function
Public Sub PesquisaItemCombo(Combo As ComboBox, Item As String)

    For x = 0 To Combo.ListCount - 1
        If Combo.ItemData(x) = Item Then
            Combo.ListIndex = x
            Exit For
        End If
    Next x
End Sub
Public Function ConverteMoeda(ByVal VALOR As String) As String

    'Substitui virgula por ponto
    
    For x = 1 To Len(VALOR)
        If Mid(VALOR, x, 1) = "," Then
            Mid(VALOR, x, 1) = "."
            Exit For
        End If
    Next x

    ConverteMoeda = VALOR
End Function
