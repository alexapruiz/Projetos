VERSION 5.00
Begin VB.Form Fechamento 
   Caption         =   "SGB - Fechamento"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDesp 
      Caption         =   "Incide Despesas Fixas Mensais"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   2715
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   330
      Left            =   4545
      TabIndex        =   9
      Top             =   690
      Width           =   435
   End
   Begin VB.TextBox TxtDataPara 
      Height          =   330
      Left            =   3360
      TabIndex        =   8
      Top             =   675
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   330
      Left            =   1605
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox TxtDataDe 
      Height          =   330
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   1140
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   405
      Left            =   1046
      TabIndex        =   3
      Top             =   1635
      Width           =   1440
   End
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "&Processar"
      Default         =   -1  'True
      Height          =   405
      Left            =   2974
      TabIndex        =   2
      Top             =   1650
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "0001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   3769
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Número Fechamento :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   942
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Até"
      Height          =   195
      Left            =   3045
      TabIndex        =   1
      Top             =   750
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "De"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   720
      Width           =   210
   End
End
Attribute VB_Name = "Fechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ULTIMO_FECHAMENTO As Integer
Private Sub CmdProcessar_Click()

    Dim Contas As New ClsContasaPagar

    On Error GoTo Processar_Fechamento_Erro

    'Verificar se as datas estão preenchidas e são datas válidas
    If (Not IsDate(TxtDataDe.Text)) Or (Not IsDate(TxtDataPara.Text)) Then
        MsgBox "As datas selecionadas para o Fechamento são inválidas", vbExclamation, "SGB"
        Exit Sub
    End If

    'Verificar se a data de fim de vigência é maior que a de inicio
    If Format(TxtDataDe.Text, "yyyymmdd") >= Format(TxtDataPara.Text, "yyyymmdd") Then
        MsgBox "Data de Fim de Vigência do Fechamento Inválida", vbExclamation, "SGB"
        Exit Sub
    End If

    If Contas.InserirFechamento(ULTIMO_FECHAMENTO + 1, TxtDataDe.Text, TxtDataPara.Text, chkDesp.Value) = False Then
        MsgBox "Não foi possível gerar o Fechamento do Período selecionado", vbExclamation, "SGB"
        Exit Sub
    Else
        MsgBox "Fechamento Criado com Sucesso", vbExclamation, "SGB"
        Unload Me
    End If

    Exit Sub

Processar_Fechamento_Erro:
    MsgBox "Erro ao Incluir os dados da Escala", vbExclamation, "SGB"
End Sub

Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Command2_Click()

    'TxtDataPara.Text = RetornaData()
End Sub
Private Sub Form_Load()

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    sSql = "select * From Fechamento where DATA_FIM_VIG = (SELECT MAX(DATA_FIM_VIG) from FECHAMENTO)"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    ULTIMO_FECHAMENTO = Rs("ID_FECHA").Value
    FECHAMENTO_ANTERIOR = Rs("DATA_FIM_VIG").Value
    
    TxtDataDe.Text = FECHAMENTO_ANTERIOR + 1
    TxtDataPara.Text = FECHAMENTO_ANTERIOR + 15

End Sub
