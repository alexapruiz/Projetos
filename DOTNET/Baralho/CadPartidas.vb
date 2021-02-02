Public Class CadPartidas
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtQtdePartidas As System.Windows.Forms.TextBox
    Friend WithEvents LblQtdePartidas As System.Windows.Forms.Label
    Friend WithEvents TxtData As System.Windows.Forms.TextBox
    Friend WithEvents LblData As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkAlex As System.Windows.Forms.CheckBox
    Friend WithEvents ChkRubia As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAdemir As System.Windows.Forms.CheckBox
    Friend WithEvents ChkNeuza As System.Windows.Forms.CheckBox
    Friend WithEvents ChkPanin As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtPontosAlex As System.Windows.Forms.TextBox
    Friend WithEvents TxtPontosRubia As System.Windows.Forms.TextBox
    Friend WithEvents TxtPontosAdemir As System.Windows.Forms.TextBox
    Friend WithEvents TxtPontosNeuza As System.Windows.Forms.TextBox
    Friend WithEvents TxtPontosPanin As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatidasPanin As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatidasNeuza As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatidasAdemir As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatidasRubia As System.Windows.Forms.TextBox
    Friend WithEvents TxtBatidasAlex As System.Windows.Forms.TextBox
    Friend WithEvents CmdInserirPartida As System.Windows.Forms.Button
    Friend WithEvents LstPartidas As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtQtdePartidas = New System.Windows.Forms.TextBox()
        Me.LblQtdePartidas = New System.Windows.Forms.Label()
        Me.TxtData = New System.Windows.Forms.TextBox()
        Me.LblData = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TxtBatidasPanin = New System.Windows.Forms.TextBox()
        Me.TxtBatidasNeuza = New System.Windows.Forms.TextBox()
        Me.TxtBatidasAdemir = New System.Windows.Forms.TextBox()
        Me.TxtBatidasRubia = New System.Windows.Forms.TextBox()
        Me.TxtBatidasAlex = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtPontosPanin = New System.Windows.Forms.TextBox()
        Me.TxtPontosNeuza = New System.Windows.Forms.TextBox()
        Me.TxtPontosAdemir = New System.Windows.Forms.TextBox()
        Me.TxtPontosRubia = New System.Windows.Forms.TextBox()
        Me.TxtPontosAlex = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ChkPanin = New System.Windows.Forms.CheckBox()
        Me.ChkNeuza = New System.Windows.Forms.CheckBox()
        Me.ChkAdemir = New System.Windows.Forms.CheckBox()
        Me.ChkRubia = New System.Windows.Forms.CheckBox()
        Me.ChkAlex = New System.Windows.Forms.CheckBox()
        Me.CmdInserirPartida = New System.Windows.Forms.Button()
        Me.LstPartidas = New System.Windows.Forms.ListBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TxtQtdePartidas, Me.LblQtdePartidas, Me.TxtData, Me.LblData})
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(502, 64)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Dados da Partida"
        '
        'TxtQtdePartidas
        '
        Me.TxtQtdePartidas.Location = New System.Drawing.Point(400, 24)
        Me.TxtQtdePartidas.MaxLength = 3
        Me.TxtQtdePartidas.Name = "TxtQtdePartidas"
        Me.TxtQtdePartidas.Size = New System.Drawing.Size(64, 22)
        Me.TxtQtdePartidas.TabIndex = 1
        Me.TxtQtdePartidas.Text = ""
        '
        'LblQtdePartidas
        '
        Me.LblQtdePartidas.Location = New System.Drawing.Point(254, 28)
        Me.LblQtdePartidas.Name = "LblQtdePartidas"
        Me.LblQtdePartidas.Size = New System.Drawing.Size(152, 24)
        Me.LblQtdePartidas.TabIndex = 6
        Me.LblQtdePartidas.Text = "Quantidade de Jogos"
        '
        'TxtData
        '
        Me.TxtData.Location = New System.Drawing.Point(126, 24)
        Me.TxtData.Name = "TxtData"
        Me.TxtData.Size = New System.Drawing.Size(104, 22)
        Me.TxtData.TabIndex = 0
        Me.TxtData.Text = ""
        '
        'LblData
        '
        Me.LblData.Location = New System.Drawing.Point(22, 28)
        Me.LblData.Name = "LblData"
        Me.LblData.Size = New System.Drawing.Size(104, 24)
        Me.LblData.TabIndex = 4
        Me.LblData.Text = "Data da Partida"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.TxtBatidasPanin, Me.TxtBatidasNeuza, Me.TxtBatidasAdemir, Me.TxtBatidasRubia, Me.TxtBatidasAlex, Me.Label3, Me.Label2, Me.TxtPontosPanin, Me.TxtPontosNeuza, Me.TxtPontosAdemir, Me.TxtPontosRubia, Me.TxtPontosAlex, Me.Label1, Me.ChkPanin, Me.ChkNeuza, Me.ChkAdemir, Me.ChkRubia, Me.ChkAlex})
        Me.GroupBox2.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(328, 216)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Dados dos Jogadores"
        '
        'TxtBatidasPanin
        '
        Me.TxtBatidasPanin.Location = New System.Drawing.Point(224, 184)
        Me.TxtBatidasPanin.Name = "TxtBatidasPanin"
        Me.TxtBatidasPanin.Size = New System.Drawing.Size(80, 22)
        Me.TxtBatidasPanin.TabIndex = 12
        Me.TxtBatidasPanin.Text = ""
        '
        'TxtBatidasNeuza
        '
        Me.TxtBatidasNeuza.Location = New System.Drawing.Point(224, 152)
        Me.TxtBatidasNeuza.Name = "TxtBatidasNeuza"
        Me.TxtBatidasNeuza.Size = New System.Drawing.Size(80, 22)
        Me.TxtBatidasNeuza.TabIndex = 11
        Me.TxtBatidasNeuza.Text = ""
        '
        'TxtBatidasAdemir
        '
        Me.TxtBatidasAdemir.Location = New System.Drawing.Point(224, 120)
        Me.TxtBatidasAdemir.Name = "TxtBatidasAdemir"
        Me.TxtBatidasAdemir.Size = New System.Drawing.Size(80, 22)
        Me.TxtBatidasAdemir.TabIndex = 10
        Me.TxtBatidasAdemir.Text = ""
        '
        'TxtBatidasRubia
        '
        Me.TxtBatidasRubia.Location = New System.Drawing.Point(224, 88)
        Me.TxtBatidasRubia.Name = "TxtBatidasRubia"
        Me.TxtBatidasRubia.Size = New System.Drawing.Size(80, 22)
        Me.TxtBatidasRubia.TabIndex = 9
        Me.TxtBatidasRubia.Text = ""
        '
        'TxtBatidasAlex
        '
        Me.TxtBatidasAlex.Location = New System.Drawing.Point(224, 56)
        Me.TxtBatidasAlex.Name = "TxtBatidasAlex"
        Me.TxtBatidasAlex.Size = New System.Drawing.Size(80, 22)
        Me.TxtBatidasAlex.TabIndex = 8
        Me.TxtBatidasAlex.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(224, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Qtde Batidas"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(120, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Qtde Pontos"
        '
        'TxtPontosPanin
        '
        Me.TxtPontosPanin.Location = New System.Drawing.Point(120, 184)
        Me.TxtPontosPanin.Name = "TxtPontosPanin"
        Me.TxtPontosPanin.Size = New System.Drawing.Size(80, 22)
        Me.TxtPontosPanin.TabIndex = 6
        Me.TxtPontosPanin.Text = ""
        '
        'TxtPontosNeuza
        '
        Me.TxtPontosNeuza.Location = New System.Drawing.Point(120, 152)
        Me.TxtPontosNeuza.Name = "TxtPontosNeuza"
        Me.TxtPontosNeuza.Size = New System.Drawing.Size(80, 22)
        Me.TxtPontosNeuza.TabIndex = 5
        Me.TxtPontosNeuza.Text = ""
        '
        'TxtPontosAdemir
        '
        Me.TxtPontosAdemir.Location = New System.Drawing.Point(120, 120)
        Me.TxtPontosAdemir.Name = "TxtPontosAdemir"
        Me.TxtPontosAdemir.Size = New System.Drawing.Size(80, 22)
        Me.TxtPontosAdemir.TabIndex = 4
        Me.TxtPontosAdemir.Text = ""
        '
        'TxtPontosRubia
        '
        Me.TxtPontosRubia.Location = New System.Drawing.Point(120, 88)
        Me.TxtPontosRubia.Name = "TxtPontosRubia"
        Me.TxtPontosRubia.Size = New System.Drawing.Size(80, 22)
        Me.TxtPontosRubia.TabIndex = 3
        Me.TxtPontosRubia.Text = ""
        '
        'TxtPontosAlex
        '
        Me.TxtPontosAlex.Location = New System.Drawing.Point(120, 56)
        Me.TxtPontosAlex.Name = "TxtPontosAlex"
        Me.TxtPontosAlex.Size = New System.Drawing.Size(80, 22)
        Me.TxtPontosAlex.TabIndex = 2
        Me.TxtPontosAlex.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Nome"
        '
        'ChkPanin
        '
        Me.ChkPanin.Location = New System.Drawing.Point(32, 187)
        Me.ChkPanin.Name = "ChkPanin"
        Me.ChkPanin.Size = New System.Drawing.Size(72, 16)
        Me.ChkPanin.TabIndex = 4
        Me.ChkPanin.Text = "Panin"
        Me.ChkPanin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ChkNeuza
        '
        Me.ChkNeuza.Location = New System.Drawing.Point(32, 155)
        Me.ChkNeuza.Name = "ChkNeuza"
        Me.ChkNeuza.Size = New System.Drawing.Size(72, 16)
        Me.ChkNeuza.TabIndex = 3
        Me.ChkNeuza.Text = "Neuza"
        Me.ChkNeuza.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ChkAdemir
        '
        Me.ChkAdemir.Location = New System.Drawing.Point(32, 123)
        Me.ChkAdemir.Name = "ChkAdemir"
        Me.ChkAdemir.Size = New System.Drawing.Size(72, 16)
        Me.ChkAdemir.TabIndex = 2
        Me.ChkAdemir.Text = "Ademir"
        Me.ChkAdemir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ChkRubia
        '
        Me.ChkRubia.Location = New System.Drawing.Point(32, 91)
        Me.ChkRubia.Name = "ChkRubia"
        Me.ChkRubia.Size = New System.Drawing.Size(72, 16)
        Me.ChkRubia.TabIndex = 1
        Me.ChkRubia.Text = "Rubia"
        Me.ChkRubia.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ChkAlex
        '
        Me.ChkAlex.Location = New System.Drawing.Point(32, 59)
        Me.ChkAlex.Name = "ChkAlex"
        Me.ChkAlex.Size = New System.Drawing.Size(72, 16)
        Me.ChkAlex.TabIndex = 0
        Me.ChkAlex.Text = "Alex"
        Me.ChkAlex.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CmdInserirPartida
        '
        Me.CmdInserirPartida.Location = New System.Drawing.Point(536, 32)
        Me.CmdInserirPartida.Name = "CmdInserirPartida"
        Me.CmdInserirPartida.Size = New System.Drawing.Size(128, 24)
        Me.CmdInserirPartida.TabIndex = 7
        Me.CmdInserirPartida.Text = "&Inserir Partida"
        '
        'LstPartidas
        '
        Me.LstPartidas.ItemHeight = 16
        Me.LstPartidas.Location = New System.Drawing.Point(344, 86)
        Me.LstPartidas.Name = "LstPartidas"
        Me.LstPartidas.Size = New System.Drawing.Size(360, 212)
        Me.LstPartidas.TabIndex = 8
        '
        'CadPartidas
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(728, 312)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.LstPartidas, Me.CmdInserirPartida, Me.GroupBox2, Me.GroupBox1})
        Me.Name = "CadPartidas"
        Me.Text = "Programa para Cálculo de Média - Cadastro de Partidas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CadPartidas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Rs As New ADODB.Recordset()

        'Carregar Lista de Partidas
        Rs.Open("SelPartidas ", Db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        Do Until Rs.EOF
            LstPartidas.Items.Add(Rs("DataPartida").Value & Space(110) & Format(Rs("IdPartida").Value, "000"))
            Rs.MoveNext()
        Loop

    End Sub

    Private Sub ChkAlex_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAlex.CheckedChanged

    End Sub

    Private Sub TxtBatidasPanin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtBatidasPanin.TextChanged

    End Sub

    Private Sub TxtBatidasAdemir_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtBatidasAdemir.TextChanged

    End Sub

    Private Sub TxtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtData.TextChanged

    End Sub

    Private Sub LstPartidas_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstPartidas.SelectedIndexChanged

    End Sub

    Private Sub LstPartidas_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstPartidas.DoubleClick

        Dim Rs As New ADODB.Recordset()

        'Limpar os campos
        TxtPontosAlex.Text = ""
        TxtBatidasAlex.Text = ""
        TxtPontosRubia.Text = ""
        TxtBatidasRubia.Text = ""
        TxtPontosAdemir.Text = ""
        TxtBatidasAdemir.Text = ""
        TxtPontosNeuza.Text = ""
        TxtBatidasNeuza.Text = ""
        TxtPontosPanin.Text = ""
        TxtBatidasPanin.Text = ""

        'Pesquisar os dados da partida selecionada
        Rs.Open("SelJogadorXPartida @IdPartida = " & Mid(LstPartidas.SelectedItem, 121, 3), Db)

        Do Until Rs.EOF
            Select Case Rs("IdJogador").Value
                Case 1
                    'Alex
                    TxtPontosAlex.Text = Rs("QtdePontos").Value
                    TxtBatidasAlex.Text = Rs("QtdeBatidas").Value
                Case 2
                    'Rubia
                    TxtPontosRubia.Text = Rs("QtdePontos").Value
                    TxtBatidasRubia.Text = Rs("QtdeBatidas").Value
                Case 3
                    'Ademir
                    TxtPontosAdemir.Text = Rs("QtdePontos").Value
                    TxtBatidasAdemir.Text = Rs("QtdeBatidas").Value
                Case 4
                    'Neuza
                    TxtPontosNeuza.Text = Rs("QtdePontos").Value
                    TxtBatidasNeuza.Text = Rs("QtdeBatidas").Value
                Case 5
                    'Panin
                    TxtPontosPanin.Text = Rs("QtdePontos").Value
                    TxtBatidasPanin.Text = Rs("QtdeBatidas").Value
            End Select
            Rs.MoveNext()
        Loop

    End Sub

    Private Sub CmdInserirPartida_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdInserirPartida.Click

        Dim Rs As New ADODB.Recordset()
        Dim Cmd As New ADODB.Command()

        'Inserir na tabela "Partidas"
        Rs.Open("inspartida '" & TxtData.Text & "'," & TxtQtdePartidas.Text, Db)

        'Inserir na tabela "JogadoresXPartidas"

        'Alex
        Cmd.CommandText = "InsJogadoresXPartidas "
        Cmd.ActiveConnection = Db
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = "InsJogadorXPartida " & Rs("Iden").Value & "," & 1 & "," & TxtPontosAlex.Text & "," & TxtBatidasAlex.Text
        Cmd.Execute()


        'Rubia
        Cmd.CommandText = "InsJogadoresXPartidas "
        Cmd.ActiveConnection = Db
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = "InsJogadorXPartida " & Rs("Iden").Value & "," & 2 & "," & TxtPontosRubia.Text & "," & TxtBatidasRubia.Text
        Cmd.Execute()

        'Ademir
        Cmd.CommandText = "InsJogadoresXPartidas "
        Cmd.ActiveConnection = Db
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = "InsJogadorXPartida " & Rs("Iden").Value & "," & 3 & "," & TxtPontosAdemir.Text & "," & TxtBatidasAdemir.Text
        Cmd.Execute()

        'Neuza
        Cmd.CommandText = "InsJogadoresXPartidas "
        Cmd.ActiveConnection = Db
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = "InsJogadorXPartida " & Rs("Iden").Value & "," & 4 & "," & TxtPontosNeuza.Text & "," & TxtBatidasNeuza.Text
        Cmd.Execute()

        'Panin
        Cmd.CommandText = "InsJogadoresXPartidas "
        Cmd.ActiveConnection = Db
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandText = "InsJogadorXPartida " & Rs("Iden").Value & "," & 5 & "," & TxtPontosPanin.Text & "," & TxtBatidasPanin.Text
        Cmd.Execute()

    End Sub
End Class
