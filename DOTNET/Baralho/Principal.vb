Public Class Principal
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
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MnuCadastros As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCadJogadores As System.Windows.Forms.MenuItem
    Friend WithEvents MnuCadPartidas As System.Windows.Forms.MenuItem
    Friend WithEvents MnuSair As System.Windows.Forms.MenuItem
    Friend WithEvents MnuInserirDados As System.Windows.Forms.MenuItem
    Friend WithEvents MnuConsulta As System.Windows.Forms.MenuItem
    Friend WithEvents MnuConsRanking As System.Windows.Forms.MenuItem
    Friend WithEvents MnuConsEstatisticas As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MnuCadastros = New System.Windows.Forms.MenuItem()
        Me.MnuCadJogadores = New System.Windows.Forms.MenuItem()
        Me.MnuCadPartidas = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MnuSair = New System.Windows.Forms.MenuItem()
        Me.MnuInserirDados = New System.Windows.Forms.MenuItem()
        Me.MnuConsulta = New System.Windows.Forms.MenuItem()
        Me.MnuConsRanking = New System.Windows.Forms.MenuItem()
        Me.MnuConsEstatisticas = New System.Windows.Forms.MenuItem()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuCadastros, Me.MenuItem5, Me.MnuSair})
        Me.MenuItem1.Text = "&Arquivo"
        '
        'MnuCadastros
        '
        Me.MnuCadastros.Index = 0
        Me.MnuCadastros.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuCadJogadores, Me.MnuCadPartidas})
        Me.MnuCadastros.Text = "Ca&dastros"
        '
        'MnuCadJogadores
        '
        Me.MnuCadJogadores.Index = 0
        Me.MnuCadJogadores.Text = "&Jogadores"
        '
        'MnuCadPartidas
        '
        Me.MnuCadPartidas.Index = 1
        Me.MnuCadPartidas.Text = "&Partidas"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "-"
        '
        'MnuSair
        '
        Me.MnuSair.Index = 2
        Me.MnuSair.Text = "Sa&ir"
        '
        'MnuInserirDados
        '
        Me.MnuInserirDados.Index = 1
        Me.MnuInserirDados.Text = "&Inserir Dados"
        '
        'MnuConsulta
        '
        Me.MnuConsulta.Index = 2
        Me.MnuConsulta.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuConsRanking, Me.MnuConsEstatisticas})
        Me.MnuConsulta.Text = "&Consulta"
        '
        'MnuConsRanking
        '
        Me.MnuConsRanking.Index = 0
        Me.MnuConsRanking.Text = "&Ranking"
        '
        'MnuConsEstatisticas
        '
        Me.MnuConsEstatisticas.Index = 1
        Me.MnuConsEstatisticas.Text = "&Estatísticas"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MnuInserirDados, Me.MnuConsulta})
        '
        'Principal
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(592, 416)
        Me.Menu = Me.MainMenu1
        Me.Name = "Principal"
        Me.Text = "Programa para Cálculo de Média"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Db.ConnectionString = "Driver=SQL Server;Server=Beatriz;Database=Baralho;"
        Db.Open()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

    End Sub

    Private Sub MnuCadJogadores_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCadJogadores.Click

        Dim Jogadores As New CadJogadores()

        Jogadores.Show()
    End Sub

    Private Sub MnuCadPartidas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuCadPartidas.Click

        Dim Partidas As New CadPartidas()

        Partidas.Show()
    End Sub

    Private Sub MnuSair_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuSair.Click

        End
    End Sub

    Private Sub MnuConsRanking_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuConsRanking.Click
        Dim Ranking As New Ranking()

        Ranking.show()
    End Sub
End Class
