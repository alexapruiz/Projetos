Attribute VB_Name = "ImpRecAvisoDif"
Public Function ImpRecAvisoDiferen�a()



Dim rsPesquisaIDBordero            As New ADODB.Recordset
Dim Selecao                        As New custodia.Selecionar
Dim vIdBordero                     As Double
Dim SemRegistros                   As Integer


Screen.MousePointer = vbHourglass

Set rsPesquisaIDBordero = g_cMainConnection.Execute(Selecao.GetRecAvisoDiferenca(Geral.DataProcessamento, CBool(0)))

If Not rsPesquisaIDBordero.EOF Then
    Principal.CrystalReport.ReportFileName = App.path & "\Reports\RelRecAvisoDiferenca.rpt"
    Principal.CrystalReport.SelectionFormula = "{AvisoDiferenca.DataOcorrencia} = " & Geral.DataProcessamento & " and {AvisoDiferenca.Gerado} = " & Val(CBool(0))
    Principal.CrystalReport.CopiesToPrinter = 1
    Principal.CrystalReport.WindowState = crptMaximized
    Principal.CrystalReport.WindowTitle = "Emiss�o do Relat�rio de Aviso de Diferen�a"
    Principal.CrystalReport.Action = 0
Else
    SemRegistros = MsgBox("Este border� n�o possui AD's.", vbExclamation, "Relat�rio de Aviso de Diferen�a")
End If

Screen.MousePointer = vbDefault

End Function
