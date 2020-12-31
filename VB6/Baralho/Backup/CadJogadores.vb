Public Class CadJogadores
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents LblNome As System.Windows.Forms.Label
    Friend WithEvents TxtNome As System.Windows.Forms.TextBox
    Friend WithEvents CmdOK As System.Windows.Forms.Button
    Friend WithEvents CmdCancelar As System.Windows.Forms.Button
    Friend WithEvents CmdExcluir As System.Windows.Forms.Button
    Friend WithEvents CmdSair As System.Windows.Forms.Button
    Friend WithEvents LstJogadores As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.LblNome = New System.Windows.Forms.Label()
        Me.TxtNome = New System.Windows.Forms.TextBox()
        Me.CmdOK = New System.Windows.Forms.Button()
        Me.CmdCancelar = New System.Windows.Forms.Button()
        Me.CmdExcluir = New System.Windows.Forms.Button()
        Me.CmdSair = New System.Windows.Forms.Button()
        Me.LstJogadores = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'LblNome
        '
        Me.LblNome.Location = New System.Drawing.Point(24, 32)
        Me.LblNome.Name = "LblNome"
        Me.LblNome.Size = New System.Drawing.Size(48, 23)
        Me.LblNome.TabIndex = 0
        Me.LblNome.Text = "Nome"
        '
        'TxtNome
        '
        Me.TxtNome.Location = New System.Drawing.Point(88, 32)
        Me.TxtNome.Name = "TxtNome"
        Me.TxtNome.Size = New System.Drawing.Size(192, 22)
        Me.TxtNome.TabIndex = 1
        Me.TxtNome.Text = ""
        '
        'CmdOK
        '
        Me.CmdOK.Location = New System.Drawing.Point(16, 392)
        Me.CmdOK.Name = "CmdOK"
        Me.CmdOK.Size = New System.Drawing.Size(88, 24)
        Me.CmdOK.TabIndex = 3
        Me.CmdOK.Text = "&OK"
        '
        'CmdCancelar
        '
        Me.CmdCancelar.Location = New System.Drawing.Point(128, 392)
        Me.CmdCancelar.Name = "CmdCancelar"
        Me.CmdCancelar.Size = New System.Drawing.Size(88, 24)
        Me.CmdCancelar.TabIndex = 4
        Me.CmdCancelar.Text = "&Cancelar"
        '
        'CmdExcluir
        '
        Me.CmdExcluir.Location = New System.Drawing.Point(240, 392)
        Me.CmdExcluir.Name = "CmdExcluir"
        Me.CmdExcluir.Size = New System.Drawing.Size(88, 24)
        Me.CmdExcluir.TabIndex = 5
        Me.CmdExcluir.Text = "&Excluir"
        '
        'CmdSair
        '
        Me.CmdSair.Location = New System.Drawing.Point(352, 392)
        Me.CmdSair.Name = "CmdSair"
        Me.CmdSair.Size = New System.Drawing.Size(88, 24)
        Me.CmdSair.TabIndex = 6
        Me.CmdSair.Text = "Sai&r"
        '
        'LstJogadores
        '
        Me.LstJogadores.ItemHeight = 16
        Me.LstJogadores.Location = New System.Drawing.Point(24, 80)
        Me.LstJogadores.Name = "LstJogadores"
        Me.LstJogadores.Size = New System.Drawing.Size(424, 276)
        Me.LstJogadores.TabIndex = 7
        '
        'CadJogadores
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(456, 432)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.LstJogadores, Me.CmdSair, Me.CmdExcluir, Me.CmdCancelar, Me.CmdOK, Me.TxtNome, Me.LblNome})
        Me.Name = "CadJogadores"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Programa de Cálculo de Média - Cadastro de Jogadores"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Preencher o grid com os Jogadores
        Call PreencheListaJogadores()

    End Sub

    Private Sub CmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCancelar.Click

        TxtNome.Text = ""
        LstJogadores.ClearSelected()
    End Sub

    Private Sub LstJogadores_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstJogadores.DoubleClick

        TxtNome.Text = LstJogadores.SelectedItem
    End Sub

    Private Sub CmdSair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSair.Click

        Me.Close()
    End Sub

    Private Sub CmdExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdExcluir.Click

        Dim Rs As New ADODB.Recordset()
        Dim Cmd As New ADODB.Command()

        'Verificar se o usuario selecionou algum item
        If LstJogadores.SelectedIndex = -1 Then
            MsgBox("Selecione um item para excluir", MsgBoxStyle.Exclamation)
        Else
            If MsgBox("Confirma a exclusão do Jogador '" & LstJogadores.SelectedItem & "' ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Cmd.ActiveConnection = Db
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
                Cmd.CommandText = "ExcluirUsuario '" & LstJogadores.SelectedItem & "'"
                Cmd.Execute()

                'Atualizar a lista de Jogadores
                Call PreencheListaJogadores()
            End If
        End If

    End Sub

    Private Sub PreencheListaJogadores()

        Dim Rs As New ADODB.Recordset()

        'Limpar a lista de Jogadores
        LstJogadores.Items.Clear()

        'Criar recordset com os dados da tabela {Jogadores}
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open("select * from jogadores", Db)

        Do Until Rs.EOF
            LstJogadores.Items.Add(Rs("Nome").Value)
            Rs.MoveNext()
        Loop
    End Sub

    Private Sub LstJogadores_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstJogadores.SelectedIndexChanged

    End Sub
End Class
