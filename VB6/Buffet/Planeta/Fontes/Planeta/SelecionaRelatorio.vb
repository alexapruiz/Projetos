Public Class SelecionaRelatorio

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Emitir relatório com os pedidos do periodo informado
        If (TxtDe1.Text <> "") And (TxtAte1.Text <> "") Then
            Relatorios.NomeRelatorio = "d:\projetos\planeta\relatorios\pedidos2.rpt"
            Relatorios.Formula = "{pedido.data_entrega} >= '02/04/1978' and "

            Relatorios.Show()
        End If
    End Sub
End Class