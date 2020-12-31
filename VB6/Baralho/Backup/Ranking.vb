Public Class Ranking
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
    Friend WithEvents LstNome As System.Windows.Forms.ListBox
    Friend WithEvents LstMedia As System.Windows.Forms.ListBox
    Friend WithEvents LstOUROS As System.Windows.Forms.ListBox
    Friend WithEvents LstMERDAS As System.Windows.Forms.ListBox
    Friend WithEvents LstBatidas As System.Windows.Forms.ListBox
    Friend WithEvents LstIndice As System.Windows.Forms.ListBox
    Friend WithEvents LstQtdPartidas As System.Windows.Forms.ListBox
    Friend WithEvents LblNome As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lstColocacao As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.LstNome = New System.Windows.Forms.ListBox()
        Me.LstMedia = New System.Windows.Forms.ListBox()
        Me.LstOUROS = New System.Windows.Forms.ListBox()
        Me.LstMERDAS = New System.Windows.Forms.ListBox()
        Me.LstBatidas = New System.Windows.Forms.ListBox()
        Me.LstIndice = New System.Windows.Forms.ListBox()
        Me.LstQtdPartidas = New System.Windows.Forms.ListBox()
        Me.LblNome = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lstColocacao = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'LstNome
        '
        Me.LstNome.BackColor = System.Drawing.SystemColors.HighlightText
        Me.LstNome.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstNome.ItemHeight = 23
        Me.LstNome.Location = New System.Drawing.Point(5, 56)
        Me.LstNome.Name = "LstNome"
        Me.LstNome.Size = New System.Drawing.Size(88, 234)
        Me.LstNome.TabIndex = 0
        '
        'LstMedia
        '
        Me.LstMedia.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstMedia.ItemHeight = 23
        Me.LstMedia.Location = New System.Drawing.Point(97, 56)
        Me.LstMedia.Name = "LstMedia"
        Me.LstMedia.Size = New System.Drawing.Size(108, 234)
        Me.LstMedia.TabIndex = 1
        '
        'LstOUROS
        '
        Me.LstOUROS.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstOUROS.ItemHeight = 23
        Me.LstOUROS.Location = New System.Drawing.Point(317, 56)
        Me.LstOUROS.Name = "LstOUROS"
        Me.LstOUROS.Size = New System.Drawing.Size(96, 234)
        Me.LstOUROS.TabIndex = 2
        '
        'LstMERDAS
        '
        Me.LstMERDAS.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstMERDAS.ItemHeight = 23
        Me.LstMERDAS.Location = New System.Drawing.Point(425, 56)
        Me.LstMERDAS.Name = "LstMERDAS"
        Me.LstMERDAS.Size = New System.Drawing.Size(88, 234)
        Me.LstMERDAS.TabIndex = 3
        '
        'LstBatidas
        '
        Me.LstBatidas.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstBatidas.ItemHeight = 23
        Me.LstBatidas.Location = New System.Drawing.Point(526, 56)
        Me.LstBatidas.Name = "LstBatidas"
        Me.LstBatidas.Size = New System.Drawing.Size(88, 234)
        Me.LstBatidas.TabIndex = 4
        '
        'LstIndice
        '
        Me.LstIndice.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstIndice.ItemHeight = 23
        Me.LstIndice.Location = New System.Drawing.Point(628, 56)
        Me.LstIndice.Name = "LstIndice"
        Me.LstIndice.Size = New System.Drawing.Size(118, 234)
        Me.LstIndice.TabIndex = 5
        '
        'LstQtdPartidas
        '
        Me.LstQtdPartidas.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstQtdPartidas.ItemHeight = 23
        Me.LstQtdPartidas.Location = New System.Drawing.Point(217, 56)
        Me.LstQtdPartidas.Name = "LstQtdPartidas"
        Me.LstQtdPartidas.Size = New System.Drawing.Size(88, 234)
        Me.LstQtdPartidas.TabIndex = 6
        '
        'LblNome
        '
        Me.LblNome.AutoSize = True
        Me.LblNome.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNome.Location = New System.Drawing.Point(13, 24)
        Me.LblNome.Name = "LblNome"
        Me.LblNome.Size = New System.Drawing.Size(74, 27)
        Me.LblNome.TabIndex = 7
        Me.LblNome.Text = "Nome"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(117, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 27)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Média"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(211, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 27)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Partidas"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(326, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 27)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Ouros"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(427, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 27)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Merdas"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(523, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(94, 27)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Batidas"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(652, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(77, 27)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Índice"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(752, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 27)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Pos."
        '
        'lstColocacao
        '
        Me.lstColocacao.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstColocacao.ItemHeight = 23
        Me.lstColocacao.Location = New System.Drawing.Point(752, 56)
        Me.lstColocacao.Name = "lstColocacao"
        Me.lstColocacao.Size = New System.Drawing.Size(48, 234)
        Me.lstColocacao.TabIndex = 15
        '
        'Ranking
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(808, 304)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lstColocacao, Me.Label7, Me.Label6, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.LblNome, Me.LstQtdPartidas, Me.LstIndice, Me.LstBatidas, Me.LstMERDAS, Me.LstOUROS, Me.LstMedia, Me.LstNome})
        Me.Name = "Ranking"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Programa para Cálculo de Média - Ranking dos Jogadores"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Ranking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call CalculaRanking()
    End Sub

    Private Sub CalculaRanking()

        Dim Rs As New ADODB.Recordset()
        Dim Indice As Double
        Dim x As Integer

        Rs.Open("CalculaDados ", Db)

        If Rs.EOF Then
            MsgBox("Erro ao Calcular Ranking")
            Exit Sub
        End If
        x = 1
        Do Until Rs.EOF
            LstNome.Items.Add(Rs("Nome").Value)
            LstMedia.Items.Add(Rs("Media").Value)
            LstQtdPartidas.Items.Add(Rs("Qtde_Partidas").Value)
            LstOUROS.Items.Add(Rs("Ouros").Value)
            LstMERDAS.Items.Add(Rs("Merdas").Value)
            LstBatidas.Items.Add(Rs("Media_Batidas").Value)
            LstIndice.Items.Add(Rs("Indice").Value)
            lstColocacao.Items.Add(x)
            x = x + 1
            Rs.MoveNext()
        Loop

    End Sub

    Private Sub DataView1_ListChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ListChangedEventArgs)

    End Sub
End Class
