VERSION 5.00
Begin VB.Form AgenciaContaAjuste 
   Caption         =   "Digitação de Agência e Conta"
   ClientHeight    =   1812
   ClientLeft      =   2904
   ClientTop       =   2880
   ClientWidth     =   5676
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1812
   ScaleWidth      =   5676
   Begin VB.Frame FrmAgenciaConta 
      Caption         =   "Informe a Agência e Conta para Ajuste"
      Height          =   1740
      Left            =   24
      TabIndex        =   4
      Top             =   24
      Width           =   5604
      Begin VB.TextBox txtVinculo 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   288
         MaxLength       =   7
         TabIndex        =   10
         Top             =   552
         Width           =   1008
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&OK"
         Height          =   324
         Left            =   3024
         TabIndex        =   2
         Top             =   1248
         Width           =   1500
      End
      Begin VB.CommandButton CmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   324
         Left            =   1032
         TabIndex        =   3
         Top             =   1248
         Width           =   1500
      End
      Begin VB.TextBox txtAgencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   372
         Left            =   1524
         MaxLength       =   4
         TabIndex        =   0
         Top             =   552
         Width           =   612
      End
      Begin VB.TextBox txtConta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   372
         Left            =   2376
         MaxLength       =   7
         TabIndex        =   1
         Top             =   552
         Width           =   960
      End
      Begin VB.Label lblVinculo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vínculo Nr."
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   300
         TabIndex        =   9
         Top             =   288
         Width           =   768
      End
      Begin VB.Label lblDiferenca 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   3504
         TabIndex        =   8
         Top             =   288
         Width           =   696
      End
      Begin VB.Label lblValorDiferenca 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3504
         TabIndex        =   7
         Top             =   552
         Width           =   1764
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1488
         TabIndex        =   6
         Top             =   288
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2400
         TabIndex        =   5
         Top             =   288
         Width           =   888
      End
   End
End
Attribute VB_Name = "AgenciaContaAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Agencia         As Integer
Dim Conta           As Long
Public m_Diferenca  As Currency
Public m_Vinculo    As Long


Private Sub CmdCancelar_Click()

    Agencia = 0
    Conta = 0
    Unload Me

End Sub
Private Sub CmdOK_Click()

    Dim sTamanho    As Integer

    'Formatar o campo 'AGENCIA'
    txtAgencia.Text = Format(txtAgencia.Text, "0000")

    'Validar Agencia e Conta
    If Val(txtAgencia.Text) <> 0 And Val(txtConta.Text) <> 0 Then
        sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
        If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
            MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
            txtAgencia.SetFocus
            Exit Sub
        Else
            Agencia = txtAgencia.Text
            Conta = txtConta.Text
            Unload Me
        End If
    Else
        MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
        txtAgencia.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
End Sub

Private Sub Form_Load()

Dim ipix As Integer
    
    ipix = 900

    If m_Diferenca = 0 Then
        Label4.Left = Label4.Left + ipix
        Label5.Left = Label5.Left + ipix
        txtAgencia.Left = txtAgencia.Left + ipix
        txtConta.Left = txtConta.Left + ipix
        
        lblDiferenca.Visible = False
        lblValorDiferenca.Visible = False
    Else
        lblValorDiferenca = FormataValor(m_Diferenca, 21)
    End If
    txtVinculo.Text = Format(m_Vinculo, "00000")

End Sub

Private Sub txtAgencia_Change()
    If Len(Trim(txtAgencia.Text)) = txtAgencia.MaxLength Then
        SendKeys "{TAB}"
        DoEvents
    End If
End Sub
Private Sub txtAgencia_GotFocus()

  txtAgencia.SelStart = 0
  txtAgencia.SelLength = txtAgencia.MaxLength

End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtConta_Change()

    If Len(Trim(txtConta.Text)) = txtConta.MaxLength Then
        SendKeys "{TAB}"
        DoEvents
    End If
End Sub
Private Sub txtConta_GotFocus()
  txtConta.SelStart = 0
  txtConta.SelLength = txtConta.MaxLength
End Sub
Private Sub txtConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub
Public Function ShowModal(ByRef pAgencia As Integer, ByRef pConta As Long) As Boolean
      
   'Inicio
    Me.Show vbModal
        
   'Alimenta variaveis p/ retorno
    pAgencia = Agencia
    pConta = Conta
    Unload Me
 
End Function


