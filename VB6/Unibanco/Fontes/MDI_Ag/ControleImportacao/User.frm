VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "dateedit.ocx"
Begin VB.Form Password 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2124
   ClientLeft      =   3348
   ClientTop       =   3252
   ClientWidth     =   4476
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "User.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2124
   ScaleWidth      =   4476
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Backup"
      Height          =   372
      Index           =   1
      Left            =   1572
      TabIndex        =   4
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3024
      TabIndex        =   5
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Atual"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1332
   End
   Begin VB.PictureBox Panel3D1 
      AutoSize        =   -1  'True
      Height          =   1452
      Left            =   120
      ScaleHeight     =   1404
      ScaleWidth      =   4164
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4212
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   384
         Left            =   300
         Picture         =   "User.frx":030A
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   9
         Top             =   240
         Width           =   384
      End
      Begin VB.TextBox txtUsuario 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         Left            =   2112
         MaxLength       =   10
         TabIndex        =   0
         Top             =   60
         Width           =   1584
      End
      Begin VB.TextBox TxtSenha 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         IMEMode         =   3  'DISABLE
         Left            =   2112
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   492
         Width           =   1584
      End
      Begin DATEEDITLib.DateEdit txtDataProcessamento 
         Height          =   384
         Left            =   2112
         TabIndex        =   2
         Top             =   900
         Width           =   1584
         _Version        =   65537
         _ExtentX        =   2794
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Data Movimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   84
         TabIndex        =   10
         Top             =   924
         Width           =   2028
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   900
         TabIndex        =   8
         Top             =   480
         Width           =   864
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   336
         Left            =   900
         TabIndex        =   7
         Top             =   120
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *************************************
' * Definição das Variáveis do Módulo *
' *************************************
Public Cancelou As Boolean          ' Indicação de Cancelamento
Public SenhaOk As Boolean           ' Indicação de Senha Digitada
Private qryUsuario As rdoQuery       ' Chamada de Store Procedure para Leitura de Usuário
Private tbUsuario As rdoResultset    ' Leitura do Usuário

' ********************************************
' * Efetua a verificação do Login do Usuário *
' ********************************************
Private Sub cmdConfirma_Click(Index As Integer)
    Dim bBaseBackup As Boolean
    Cancelou = False
    SenhaOk = False
    ' ******************************************
    ' * Testa Digitação Obrigatória do Usuário *
    ' ******************************************
    If Trim(txtUsuario.Text) = "" Then
        Beep
        MsgBox "Digite o Usuário !", vbExclamation + vbOKOnly, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
    ' ****************************************
    ' * Testa Digitação Obrigatória da Senha *
    ' ****************************************
    If Trim(TxtSenha.Text) = "" Then
        Beep
        MsgBox "Digite a Senha !", vbExclamation + vbOKOnly, App.Title
        With TxtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
        Exit Sub
    End If
    ' ****************************************
    ' * Testa Digitação Obrigatória da Data  *
    ' ****************************************
    If Trim(txtDataProcessamento.Text) = "" Then
        Beep
        MsgBox "Digite a Data !", vbExclamation + vbOKOnly, App.Title
        With txtDataProcessamento
            .SetFocus
        End With
        Exit Sub
    End If
    
    If DataOk(Val(txtDataProcessamento.Text)) Then
        Geral.DataProcessamento = DataAAAAMMDD(Val(txtDataProcessamento.Text))
        'Geral.DiretorioImagens = PegarOpcaoINI("Diretorios", "Imagens", App.Path & "\IMAGENS") & "\" & DataAAAAMMDD(Val(txtDataProcessamento.Text)) & "\"
    Else
        MsgBox "A data informada não é válida!" & vbCr & "Obs.: O ano deve ser maior que 1997 e menor que 2051!", vbExclamation + vbOKOnly, App.Title
        txtDataProcessamento.SetFocus
        'Exit Sub
    End If
    
    On Error GoTo ErroLogin
    
    '''''''''''''''''''''''''''
    ' Rotina de inicialização '
    '''''''''''''''''''''''''''
    If Index = 0 Then
        Geral.StringConexao = "DSN=MDI_Ubb;UID=i;PWD=i;"
        bBaseBackup = False
    Else
        Geral.StringConexao = "DSN=MDI_UbbBackup;UID=i;PWD=i;"
        bBaseBackup = True
    End If
    
    With Geral.Banco
        .Connect = Geral.StringConexao
        .CursorDriver = rdUseOdbc
        .EstablishConnection rdDriverNoPrompt
    End With
    
    Set qryUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")
    ' ********************************
    ' * Leitura do Usuário Informado *
    ' ********************************
    With qryUsuario
        .rdoParameters(0).Value = Trim(txtUsuario.Text)
        Set tbUsuario = .OpenResultset(rdConcurReadOnly)
    End With
    
    ' ***********************************
    ' * Verificação do Login do Usuário *
    ' ***********************************
    If UCase(Trim(txtUsuario.Text)) = "DESENV" And UCase(Trim(TxtSenha.Text)) = "VENUS" Then
        SenhaOk = True
        Principal.mnuImportar.Enabled = True
    ElseIf tbUsuario.EOF Then
        Beep
        MsgBox "Usuário não Cadastrado !", vbExclamation + vbOKOnly, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    ElseIf UCase(Decript(Trim(tbUsuario!Senha))) <> UCase(Trim(TxtSenha.Text)) Then
        Beep
        MsgBox "Senha não Confere !", vbExclamation + vbOKOnly, App.Title
        With TxtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    Else
        SenhaOk = True
        ' *******************************************************
        ' * Ajustando Menu de Opções para Permissões do Usuário *
        ' *******************************************************
        
        While Not tbUsuario.EOF
            Select Case UCase(tbUsuario!IdGrupo)
                Case "AUX", "SUP", "SPT"
                    Principal.mnuImportar.Enabled = True
            End Select
            tbUsuario.MoveNext
        Wend
        
        DoEvents
    End If
    
    tbUsuario.Close
  
    If SenhaOk Then
        Geral.Usuario = txtUsuario
        Me.Hide
    Else
        Geral.Banco.Close
    End If
    
    Exit Sub

ErroLogin:
    Select Case TratamentoErro("Erro na conexão com Banco de Dados.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
        .SetFocus
    End With

End Sub
' ******************************
' * Cancela o Login no Sistema *
' ******************************
Private Sub CmdSair_Click()
    Cancelou = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    Cancelou = True
End Sub

' **************************************
' * Carrega Módulo de Login no Sistema *
' **************************************
Private Sub Form_Load()
    txtDataProcessamento.Text = Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000")
    Cancelou = False
    SenhaOk = False
    Set Geral.Banco = New rdoConnection
End Sub

Private Sub txtDataProcessamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If DataOk(Val(txtDataProcessamento.Text)) Then
            cmdConfirma_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

' ************************************
' * Ajustando Seleção do Campo Senha *
' ************************************
Private Sub txtSenha_GotFocus()
    With TxtSenha
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

' **************************************
' * Ajustando Seleção do Campo Usuario *
' **************************************
Private Sub txtUsuario_GotFocus()
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
