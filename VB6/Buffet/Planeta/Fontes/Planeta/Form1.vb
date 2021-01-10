Public Class Principal

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        TelaPedidos.ShowDialog()
    End Sub

    Private Sub Principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Redimensiona e centraliza o form na tela
        Me.Width = 1024
        Me.Height = 768
        Me.CenterToScreen()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TelaClientes.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        SelecionaRelatorio.ShowDialog()
    End Sub
End Class
