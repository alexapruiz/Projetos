Imports MySql.Data.MySqlClient

Public Class TelaPedidos

    Inherits System.Windows.Forms.Form

    Public myCommand As New MySqlCommand
    Public myAdapter As New MySqlDataAdapter
    Public myData As New DataTable
    Friend WithEvents CmdLimpaCampos As System.Windows.Forms.Button
    Public SQL As String
    Private Sub TelaPedidos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Redimensiona e centraliza o form na tela
        Me.Width = 980
        Me.Height = 659
        Me.CenterToScreen()

        'Dim conn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset
        Dim Rs2 As New ADODB.Recordset
        Dim Cli As New Clientes
        Dim Ped As New Pedidos

        'Consulta os pedidos
        LstPedidos.Items.Clear()
        Ped.ConsultaPedido(0, Rs)
        While Not Rs.EOF
            LstPedidos.Items.Add(Format(Rs("codigo").Value, "000") & "  -  " & Rs("data_entrega").Value & " - " & Rs("hora_entrega").Value)
            Rs.MoveNext()
        End While

        'Consulta os clientes
        LstClientes.Items.Clear()
        Cli.LeClientes(0, Rs2)
        While Not Rs2.EOF
            LstClientes.Items.Add(Format(Rs2("codigo").Value, "000") & " - " & Rs2("nome").Value)
            Rs2.MoveNext()
        End While
    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        Call preencheList1(1, False)
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

        Call preencheList1(2, False)
    End Sub
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        Call preencheList1(3, False)
    End Sub
    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged

        Call preencheList1(4, True)
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Me.Close()
    End Sub
    Private Sub preencheList1(ByVal codigo As Integer, ByVal mostra_tema As Integer)

        Dim Rs As New ADODB.Recordset
        Dim it As New Item

        LstItem1.Items.Clear()
        it.ConsultaItem(0, codigo, Rs)

        Do Until Rs.EOF
            LstItem1.Items.Add(Format(Rs("Codigo").Value, "000") & "    -   " & Rs("descricao").Value)

            Rs.MoveNext()
        Loop
        TxtTema.Visible = mostra_tema
        LblTema.Visible = mostra_tema
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (Val(TxtQtde.Text) > 0) Then
            If (LstItem1.SelectedIndex <> -1) Then
                LstItem2.Items.Add(Format(Val(TxtQtde.Text), "000") & "  -  " & Mid(LstItem1.Items(LstItem1.SelectedIndex), 12, Len(LstItem1.Items(LstItem1.SelectedIndex))) & Space(200) & Mid(LstItem1.Items(LstItem1.SelectedIndex), 1, 3))
            Else
                MsgBox("Informe o item !!!", MsgBoxStyle.Exclamation)
            End If
        Else
            MsgBox("Informe a Quantidade !!!", MsgBoxStyle.Exclamation)
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'INCLUIR O PEDIDO NA BASE DE DADOS
        Dim Ped As New Pedidos
        Dim Rs As New ADODB.Recordset
        Dim Rs2 As New ADODB.Recordset
        Dim Cli As New Clientes

        'Verificar se todos os itens foram preenchidos
        If (IsDate(TxtDataPedido.Text)) And (TxtHoraPedido.Text <> "") And (LstItem2.Items.Count > 0) Then

            'Incluir o pedido
            Ped.Incluir(TxtDataPedido.Text, TxtHoraPedido.Text, Mid(LstClientes.Items(LstClientes.SelectedIndex), 1, 3))

            'Recuperar o codigo do pedido
            Ped.ConsultaPedido(-1, Rs)
            TxtPedido.Text = Rs(0).Value

            'Incluir os itens do pedido
            For x = 0 To LstItem2.Items.Count - 1
                Ped.IncluirItens(TxtPedido.Text, Mid(LstItem2.Items(x), Len(LstItem2.Items(x)) - 2, 3), Mid(LstItem2.Items(x), 1, 3), "")
            Next

            'Atualizar a lista de pedidos
            LstPedidos.Items.Clear()
            Ped.ConsultaPedido(0, Rs2)
            While Not Rs2.EOF
                LstPedidos.Items.Add(Format(Rs2("codigo").Value, "000") & "  -  " & Rs2("data_entrega").Value & " - " & Rs2("hora_entrega").Value)
                Rs2.MoveNext()
            End While

            MsgBox("Foi gerado o pedido nº:" & TxtPedido.Text)

            'Limpa todos os campos
            LstPedidos.ClearSelected()
            LstClientes.ClearSelected()
            LstItem1.ClearSelected()
            LstItem2.Items.Clear()
            TxtPedido.Text = ""
            TxtDataPedido.Text = ""
            TxtHoraPedido.Text = ""
            TxtQtde.Text = ""
            TxtTema.Text = ""
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        LstItem2.Items.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub
End Class