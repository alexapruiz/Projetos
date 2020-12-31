Public Class TelaRelatorios

    Private Sub TelaRelatorios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Redimensiona e centraliza o form na tela
        Me.Width = 980
        Me.Height = 659
        Me.CenterToScreen()
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class