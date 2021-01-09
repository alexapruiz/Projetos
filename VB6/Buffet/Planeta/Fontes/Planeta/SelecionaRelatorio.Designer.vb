<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelecionaRelatorio
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtDe1 = New System.Windows.Forms.TextBox
        Me.LblAte1 = New System.Windows.Forms.Label
        Me.TxtAte1 = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TxtAte1)
        Me.Panel1.Controls.Add(Me.LblAte1)
        Me.Panel1.Controls.Add(Me.TxtDe1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Location = New System.Drawing.Point(17, 46)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(435, 54)
        Me.Panel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pedidos da Semana"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(342, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Visualizar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(21, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "De"
        '
        'TxtDe1
        '
        Me.TxtDe1.Location = New System.Drawing.Point(55, 14)
        Me.TxtDe1.Name = "TxtDe1"
        Me.TxtDe1.Size = New System.Drawing.Size(100, 20)
        Me.TxtDe1.TabIndex = 2
        '
        'LblAte1
        '
        Me.LblAte1.AutoSize = True
        Me.LblAte1.Location = New System.Drawing.Point(182, 17)
        Me.LblAte1.Name = "LblAte1"
        Me.LblAte1.Size = New System.Drawing.Size(23, 13)
        Me.LblAte1.TabIndex = 3
        Me.LblAte1.Text = "Até"
        '
        'TxtAte1
        '
        Me.TxtAte1.Location = New System.Drawing.Point(211, 14)
        Me.TxtAte1.Name = "TxtAte1"
        Me.TxtAte1.Size = New System.Drawing.Size(100, 20)
        Me.TxtAte1.TabIndex = 4
        '
        'SelecionaRelatorio
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(869, 530)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "SelecionaRelatorio"
        Me.Text = "Seleciona o Relatorio"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TxtAte1 As System.Windows.Forms.TextBox
    Friend WithEvents LblAte1 As System.Windows.Forms.Label
    Friend WithEvents TxtDe1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
