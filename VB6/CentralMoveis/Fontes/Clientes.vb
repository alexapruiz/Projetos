'Imports MySql.Data.MySqlClient
Public Class Clientes
    Inherits System.Windows.Forms.Form
    Friend WithEvents CmdCancelar As System.Windows.Forms.Button
    Friend WithEvents CmdIncluir As System.Windows.Forms.Button
    Friend WithEvents CmdExcluir As System.Windows.Forms.Button
    Friend WithEvents CmdConfirmar As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtNome As System.Windows.Forms.TextBox
    Friend WithEvents TxtCPF As System.Windows.Forms.TextBox
    Friend WithEvents TxtCNPJ As System.Windows.Forms.TextBox
    Friend WithEvents TxtRG As System.Windows.Forms.TextBox
    Friend WithEvents TxtIE As System.Windows.Forms.TextBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGridView

    'Variaveis do banco de dados
    Public conn As New MySqlConnection
    Public myCommand As New MySqlCommand
    Public myAdapter As New MySqlDataAdapter
    Public myData As New DataTable
    Friend WithEvents CmdLimpaCampos As System.Windows.Forms.Button
    Dim SQL As String
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CmdCancelar = New System.Windows.Forms.Button
        Me.CmdIncluir = New System.Windows.Forms.Button
        Me.CmdExcluir = New System.Windows.Forms.Button
        Me.CmdConfirmar = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtNome = New System.Windows.Forms.TextBox
        Me.TxtCPF = New System.Windows.Forms.TextBox
        Me.TxtCNPJ = New System.Windows.Forms.TextBox
        Me.TxtRG = New System.Windows.Forms.TextBox
        Me.TxtIE = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGridView
        Me.CmdLimpaCampos = New System.Windows.Forms.Button
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmdCancelar
        '
        Me.CmdCancelar.Location = New System.Drawing.Point(12, 589)
        Me.CmdCancelar.Name = "CmdCancelar"
        Me.CmdCancelar.Size = New System.Drawing.Size(75, 23)
        Me.CmdCancelar.TabIndex = 1
        Me.CmdCancelar.Text = "Cancelar"
        Me.CmdCancelar.UseVisualStyleBackColor = True
        '
        'CmdIncluir
        '
        Me.CmdIncluir.Location = New System.Drawing.Point(311, 589)
        Me.CmdIncluir.Name = "CmdIncluir"
        Me.CmdIncluir.Size = New System.Drawing.Size(75, 23)
        Me.CmdIncluir.TabIndex = 2
        Me.CmdIncluir.Text = "Incluir"
        Me.CmdIncluir.UseVisualStyleBackColor = True
        '
        'CmdExcluir
        '
        Me.CmdExcluir.Location = New System.Drawing.Point(610, 589)
        Me.CmdExcluir.Name = "CmdExcluir"
        Me.CmdExcluir.Size = New System.Drawing.Size(75, 23)
        Me.CmdExcluir.TabIndex = 3
        Me.CmdExcluir.Text = "Excluir"
        Me.CmdExcluir.UseVisualStyleBackColor = True
        '
        'CmdConfirmar
        '
        Me.CmdConfirmar.Location = New System.Drawing.Point(910, 589)
        Me.CmdConfirmar.Name = "CmdConfirmar"
        Me.CmdConfirmar.Size = New System.Drawing.Size(75, 23)
        Me.CmdConfirmar.TabIndex = 4
        Me.CmdConfirmar.Text = "Confirmar"
        Me.CmdConfirmar.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Nome"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(623, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "CPF"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(809, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "CNPJ"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(21, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(23, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "RG"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(264, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "IE"
        '
        'TxtNome
        '
        Me.TxtNome.Location = New System.Drawing.Point(50, 18)
        Me.TxtNome.MaxLength = 100
        Me.TxtNome.Name = "TxtNome"
        Me.TxtNome.Size = New System.Drawing.Size(435, 20)
        Me.TxtNome.TabIndex = 10
        '
        'TxtCPF
        '
        Me.TxtCPF.Location = New System.Drawing.Point(656, 21)
        Me.TxtCPF.MaxLength = 11
        Me.TxtCPF.Name = "TxtCPF"
        Me.TxtCPF.Size = New System.Drawing.Size(122, 20)
        Me.TxtCPF.TabIndex = 11
        '
        'TxtCNPJ
        '
        Me.TxtCNPJ.Location = New System.Drawing.Point(849, 18)
        Me.TxtCNPJ.MaxLength = 14
        Me.TxtCNPJ.Name = "TxtCNPJ"
        Me.TxtCNPJ.Size = New System.Drawing.Size(136, 20)
        Me.TxtCNPJ.TabIndex = 12
        '
        'TxtRG
        '
        Me.TxtRG.Location = New System.Drawing.Point(50, 44)
        Me.TxtRG.MaxLength = 10
        Me.TxtRG.Name = "TxtRG"
        Me.TxtRG.Size = New System.Drawing.Size(134, 20)
        Me.TxtRG.TabIndex = 14
        '
        'TxtIE
        '
        Me.TxtIE.Location = New System.Drawing.Point(287, 44)
        Me.TxtIE.MaxLength = 12
        Me.TxtIE.Name = "TxtIE"
        Me.TxtIE.Size = New System.Drawing.Size(158, 20)
        Me.TxtIE.TabIndex = 15
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(491, 16)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(30, 23)
        Me.Button5.TabIndex = 16
        Me.Button5.Text = "..."
        Me.Button5.UseVisualStyleBackColor = True
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowUserToAddRows = False
        Me.DataGrid1.AllowUserToDeleteRows = False
        Me.DataGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGrid1.Location = New System.Drawing.Point(12, 70)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGrid1.Size = New System.Drawing.Size(973, 513)
        Me.DataGrid1.TabIndex = 18
        '
        'CmdLimpaCampos
        '
        Me.CmdLimpaCampos.Location = New System.Drawing.Point(491, 41)
        Me.CmdLimpaCampos.Name = "CmdLimpaCampos"
        Me.CmdLimpaCampos.Size = New System.Drawing.Size(136, 23)
        Me.CmdLimpaCampos.TabIndex = 19
        Me.CmdLimpaCampos.Text = "Limpar Pesquisa"
        Me.CmdLimpaCampos.UseVisualStyleBackColor = True
        '
        'Clientes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1003, 619)
        Me.Controls.Add(Me.CmdLimpaCampos)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.TxtIE)
        Me.Controls.Add(Me.TxtRG)
        Me.Controls.Add(Me.TxtCNPJ)
        Me.Controls.Add(Me.TxtCPF)
        Me.Controls.Add(Me.TxtNome)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmdConfirmar)
        Me.Controls.Add(Me.CmdExcluir)
        Me.Controls.Add(Me.CmdIncluir)
        Me.Controls.Add(Me.CmdCancelar)
        Me.Name = "Clientes"
        Me.Text = "Cadastro de Clientes"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
    Private Sub ConsultaClientes(ByVal SQL As String)

        Try
            myData.Clear()
            myCommand.Connection = conn
            myCommand.CommandText = SQL
            myAdapter.SelectCommand = myCommand
            myAdapter.Fill(myData)
            DataGrid1.DataSource = myData
        Catch myerror As MySqlException
            MsgBox("Erro ao acessar banco de dados: " & myerror.Message)
        End Try
    End Sub
    Private Sub Clientes_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Call PreparaGrid()
    End Sub
    Private Sub PreparaGrid()

        DataGrid1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        DataGrid1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        DataGrid1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        DataGrid1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        DataGrid1.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        DataGrid1.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Redimensiona e centraliza o form na tela
        Me.Width = 1024
        Me.Height = 768
        Me.CenterToScreen()

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "server=localhost;user id=root;password=bia1701;database=CentralMoveis"

        'Abre o banco de dados
        Try
            conn.Open()
        Catch myerror As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerror.Message)
        Finally
            conn.Dispose()
        End Try

        SQL = "select * from clientes"
        Call ConsultaClientes("select * from clientes")
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        If TxtNome.Text <> "" Then
            Call ConsultaClientes("select * from clientes where nome like '" & TxtNome.Text & "%'")
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Call ConsultaClientes("select * from clientes")
    End Sub
    Private Sub CmdLimpaCampos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdLimpaCampos.Click

        Call LimpaCamposDetalhe()
        Call ConsultaClientes("select * from clientes")
    End Sub
    Private Sub LimpaCamposDetalhe()

        TxtNome.Text = ""
        TxtCPF.Text = ""
        TxtCNPJ.Text = ""
        TxtRG.Text = ""
        TxtIE.Text = ""
        TxtNome.Focus()
    End Sub
    Private Sub CmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCancelar.Click

        Call LimpaCamposDetalhe()
    End Sub
    Private Sub CmdIncluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdIncluir.Click

        'Verificar se os campos estão preenchidos
        If VerificaCamposObrigatorios() = True Then
            'Adicionar o novo registro na base de dados
            Call AdicionaRegistro()
            'Refazer o grid para que a operação faça efeito
            Call ConsultaClientes("select * from clientes")
        End If

    End Sub
    Private Sub CmdExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdExcluir.Click

        'Solicitar confirmação da exclusão do registro

        'Excluir o registro na base de dados

        'Refazer o grid para que a operação faça efeito

    End Sub
    Private Function VerificaCamposObrigatorios() As Integer

        'Verificar os campos obrigatórios
        VerificaCamposObrigatorios = True
    End Function
    Private Sub AdicionaRegistro()

        'Incluir o registor no banco de dados
    End Sub
End Class