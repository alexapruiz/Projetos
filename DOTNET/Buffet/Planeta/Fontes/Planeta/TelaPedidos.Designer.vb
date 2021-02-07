<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TelaPedidos
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.LstClientes = New System.Windows.Forms.ListBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.LstItem1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtQtde = New System.Windows.Forms.TextBox
        Me.TxtTema = New System.Windows.Forms.TextBox
        Me.LblTema = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.LstItem2 = New System.Windows.Forms.ListBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.LstPedidos = New System.Windows.Forms.ListBox
        Me.TxtPedido = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtDataPedido = New System.Windows.Forms.TextBox
        Me.TxtHoraPedido = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 287)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Clientes"
        '
        'LstClientes
        '
        Me.LstClientes.FormattingEnabled = True
        Me.LstClientes.Location = New System.Drawing.Point(73, 287)
        Me.LstClientes.Name = "LstClientes"
        Me.LstClientes.Size = New System.Drawing.Size(176, 69)
        Me.LstClientes.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(461, 498)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(47, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = ">>"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(461, 536)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(47, 23)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "<<"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'LstItem1
        '
        Me.LstItem1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.LstItem1.FormattingEnabled = True
        Me.LstItem1.Location = New System.Drawing.Point(27, 425)
        Me.LstItem1.Name = "LstItem1"
        Me.LstItem1.Size = New System.Drawing.Size(317, 134)
        Me.LstItem1.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(418, 428)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Qtde"
        '
        'TxtQtde
        '
        Me.TxtQtde.Location = New System.Drawing.Point(451, 425)
        Me.TxtQtde.Name = "TxtQtde"
        Me.TxtQtde.Size = New System.Drawing.Size(67, 20)
        Me.TxtQtde.TabIndex = 8
        '
        'TxtTema
        '
        Me.TxtTema.Location = New System.Drawing.Point(451, 459)
        Me.TxtTema.Name = "TxtTema"
        Me.TxtTema.Size = New System.Drawing.Size(67, 20)
        Me.TxtTema.TabIndex = 10
        Me.TxtTema.Visible = False
        '
        'LblTema
        '
        Me.LblTema.AutoSize = True
        Me.LblTema.Location = New System.Drawing.Point(414, 462)
        Me.LblTema.Name = "LblTema"
        Me.LblTema.Size = New System.Drawing.Size(34, 13)
        Me.LblTema.TabIndex = 9
        Me.LblTema.Text = "Tema"
        Me.LblTema.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.RadioButton4)
        Me.Panel1.Controls.Add(Me.RadioButton3)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Location = New System.Drawing.Point(27, 362)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(916, 35)
        Me.Panel1.TabIndex = 15
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(830, 7)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(51, 17)
        Me.RadioButton4.TabIndex = 18
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "Bolos"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(593, 7)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(56, 17)
        Me.RadioButton3.TabIndex = 17
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Doces"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(300, 7)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(112, 17)
        Me.RadioButton2.TabIndex = 16
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Salgados Assados"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(22, 7)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(97, 17)
        Me.RadioButton1.TabIndex = 15
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Salgados Fritos"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(586, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(124, 29)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "PEDIDO: "
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(827, 575)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(115, 36)
        Me.Button3.TabIndex = 17
        Me.Button3.Text = "Incluir Pedido"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(27, 575)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(115, 36)
        Me.Button4.TabIndex = 18
        Me.Button4.Text = "Sai&r"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'LstItem2
        '
        Me.LstItem2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.LstItem2.FormattingEnabled = True
        Me.LstItem2.Location = New System.Drawing.Point(625, 425)
        Me.LstItem2.Name = "LstItem2"
        Me.LstItem2.Size = New System.Drawing.Size(317, 134)
        Me.LstItem2.TabIndex = 19
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(623, 409)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Qtde"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(659, 409)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Descrição"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(24, 409)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 13)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Código"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(70, 409)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 13)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "Descrição"
        '
        'LstPedidos
        '
        Me.LstPedidos.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstPedidos.FormattingEnabled = True
        Me.LstPedidos.ItemHeight = 24
        Me.LstPedidos.Location = New System.Drawing.Point(26, 26)
        Me.LstPedidos.Name = "LstPedidos"
        Me.LstPedidos.Size = New System.Drawing.Size(491, 196)
        Me.LstPedidos.TabIndex = 25
        '
        'TxtPedido
        '
        Me.TxtPedido.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPedido.ForeColor = System.Drawing.Color.Blue
        Me.TxtPedido.Location = New System.Drawing.Point(716, 26)
        Me.TxtPedido.Name = "TxtPedido"
        Me.TxtPedido.Size = New System.Drawing.Size(156, 35)
        Me.TxtPedido.TabIndex = 27
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(273, 290)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(81, 13)
        Me.Label8.TabIndex = 28
        Me.Label8.Text = "Data do Pedido"
        '
        'TxtDataPedido
        '
        Me.TxtDataPedido.Location = New System.Drawing.Point(361, 287)
        Me.TxtDataPedido.MaxLength = 10
        Me.TxtDataPedido.Name = "TxtDataPedido"
        Me.TxtDataPedido.Size = New System.Drawing.Size(100, 20)
        Me.TxtDataPedido.TabIndex = 29
        '
        'TxtHoraPedido
        '
        Me.TxtHoraPedido.Location = New System.Drawing.Point(361, 317)
        Me.TxtHoraPedido.MaxLength = 5
        Me.TxtHoraPedido.Name = "TxtHoraPedido"
        Me.TxtHoraPedido.Size = New System.Drawing.Size(100, 20)
        Me.TxtHoraPedido.TabIndex = 31
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(273, 321)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(81, 13)
        Me.Label9.TabIndex = 30
        Me.Label9.Text = "Hora do Pedido"
        '
        'TelaPedidos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(964, 623)
        Me.Controls.Add(Me.TxtHoraPedido)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TxtDataPedido)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TxtPedido)
        Me.Controls.Add(Me.LstPedidos)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LstItem2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TxtTema)
        Me.Controls.Add(Me.LblTema)
        Me.Controls.Add(Me.TxtQtde)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LstItem1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.LstClientes)
        Me.Controls.Add(Me.Label1)
        Me.Name = "TelaPedidos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pedidos"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LstClientes As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents LstItem1 As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtQtde As System.Windows.Forms.TextBox
    Friend WithEvents TxtTema As System.Windows.Forms.TextBox
    Friend WithEvents LblTema As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents LstItem2 As System.Windows.Forms.ListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents LstPedidos As System.Windows.Forms.ListBox
    Friend WithEvents TxtPedido As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtDataPedido As System.Windows.Forms.TextBox
    Friend WithEvents TxtHoraPedido As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
