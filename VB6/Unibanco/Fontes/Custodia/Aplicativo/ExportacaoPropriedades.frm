VERSION 5.00
Begin VB.Form ExportacaoPropriedades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5064
   ClientLeft      =   2448
   ClientTop       =   2028
   ClientWidth     =   4764
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5064
   ScaleWidth      =   4764
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPropriedadeListView 
      Height          =   4788
      Left            =   168
      TabIndex        =   0
      Top             =   72
      Width           =   4428
      Begin VB.Frame fraTamanho 
         Caption         =   "Máximo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   684
         Left            =   648
         TabIndex        =   11
         Top             =   336
         Width           =   3108
         Begin VB.TextBox txtTamanho 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   2256
            MaxLength       =   3
            TabIndex        =   12
            Top             =   264
            Width           =   636
         End
         Begin VB.Label lblTamanho 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tamanho do campo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   192
            TabIndex        =   13
            Top             =   264
            Width           =   1908
         End
      End
      Begin VB.Frame fraAlinha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alinhamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1308
         Left            =   648
         TabIndex        =   7
         Top             =   1200
         Width           =   3108
         Begin VB.OptionButton optAlinha_Sem 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sem alinhamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   720
            TabIndex        =   10
            Top             =   312
            Width           =   1812
         End
         Begin VB.OptionButton optAlinha_Esquerdo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "À Esquerda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   720
            TabIndex        =   9
            Top             =   600
            Width           =   1812
         End
         Begin VB.OptionButton optAlinha_Direito 
            BackColor       =   &H00C0C0C0&
            Caption         =   "À Direita"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Top             =   888
            Width           =   1812
         End
      End
      Begin VB.Frame fraZeros 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preenchimento com zeros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1284
         Left            =   648
         TabIndex        =   3
         Top             =   2712
         Width           =   3108
         Begin VB.OptionButton optZeros_Direita 
            BackColor       =   &H00C0C0C0&
            Caption         =   "À Direita"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   696
            TabIndex        =   6
            Top             =   864
            Width           =   1812
         End
         Begin VB.OptionButton optZeros_Esquerda 
            BackColor       =   &H00C0C0C0&
            Caption         =   "À Esquerda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   696
            TabIndex        =   5
            Top             =   576
            Width           =   1812
         End
         Begin VB.OptionButton optZeros_Sem 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Não preencher"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   696
            TabIndex        =   4
            Top             =   288
            Width           =   1812
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   324
         Left            =   2304
         TabIndex        =   2
         Top             =   4200
         Width           =   1452
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   324
         Left            =   672
         TabIndex        =   1
         Top             =   4200
         Width           =   1452
      End
   End
End
Attribute VB_Name = "ExportacaoPropriedades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private FormChamador     As Form
Private Cancel           As Boolean

Private Sub cmdConfirmar_Click()

     If Not (Me.optAlinha_Sem Xor Me.optAlinha_Esquerdo Xor Me.optAlinha_Direito) Then
          Beep
          MsgBox "Favor selecionar um dos tipos de alinhamentos", vbInformation, Me.Caption
          Exit Sub
     End If
     
     If Not (Me.optZeros_Sem Xor Me.optZeros_Esquerda Xor Me.optZeros_Direita) Then
          Beep
          MsgBox "Favor selecionar um dos tipos de preenchimento com zeros", vbInformation, Me.Caption
          Exit Sub
     End If
     
     If Val(txtTamanho.Text) > Val(txtTamanho.Tag) Then
          Beep
          MsgBox "Tamanho excedeu o limite máximo do campo que é de ( " & CStr(txtTamanho.Tag) & " )", vbInformation, Me.Caption
          txtTamanho.SelStart = 0
          txtTamanho.SelLength = txtTamanho.MaxLength
          txtTamanho.SetFocus
          Exit Sub
     End If
     
     If Val(txtTamanho.Text) = 0 Then
          Beep
          MsgBox "Tamanho do campo deve estar entre 1 e " & CStr(txtTamanho.Tag), vbInformation, Me.Caption
          txtTamanho.SetFocus
          Exit Sub
     End If
     
     Cancel = False
     Me.Hide
     
End Sub

Private Sub cmdSair_Click()
     
     Me.Hide
     
End Sub

Private Sub Form_Activate()

     Dim iCount As Integer
     
     Cancel = True
     
     Me.Top = (Screen.Height - Me.Height) / 2
     Me.Left = (Screen.Width - Me.Width) / 2

     If FormChamador.cmbDelimitador.ListIndex <> 0 Then
          optAlinha_Sem.Value = True
          optAlinha_Direito.Enabled = False
          optAlinha_Esquerdo.Enabled = False
     Else
          optAlinha_Sem.Enabled = False
          optAlinha_Direito.Enabled = True
          optAlinha_Esquerdo.Enabled = True
     End If
     
     With txtTamanho
          .Text = FormChamador.ListView.SelectedItem.SubItems(FormChamador.m_Col_Tamanho)
          .SelStart = 0
          .SelLength = .MaxLength
     End With
     
End Sub

Public Sub SetForm(frmFormName As Form)

     Set FormChamador = frmFormName

End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
     
     If KeyCode = vbKeyEscape Then Me.Hide
     
End Sub
Public Function GetCancelou() As Boolean

     GetCancelou = Cancel
     
End Function

Private Sub txtTamanho_KeyPress(KeyAscii As Integer)

     If KeyAscii = vbKeyEscape Then Exit Sub
     If KeyAscii = vbKeyReturn Then
          cmdConfirmar.SetFocus
          Exit Sub
     End If
     
     If KeyAscii < 48 Or KeyAscii > 57 Then Exit Sub
     
     If Len(txtTamanho) >= 3 Then
          KeyAscii = 0
          Beep
          MsgBox "Número máximo permitido é de 3 dígitos", vbInformation, Me.Caption
     End If

End Sub
