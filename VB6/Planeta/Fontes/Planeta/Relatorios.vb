Public Class Relatorios
    Public NomeRelatorio As String
    Public Formula As String

    Private Sub Relatorios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        CrystalReportViewer1.ReportSource = NomeRelatorio

        CrystalReportViewer1.SelectionFormula = Formula

        CrystalReportViewer1.Refresh()
    End Sub
End Class