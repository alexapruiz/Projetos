VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Reserva 
   Caption         =   "SGB - Reserva de Datas"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "Limpar Campos"
      Height          =   405
      Left            =   4743
      TabIndex        =   9
      Top             =   4875
      Width           =   1440
   End
   Begin VB.ComboBox CboFiltro 
      Height          =   315
      Left            =   6105
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   120
      Width           =   2310
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3195
      Left            =   60
      TabIndex        =   18
      Top             =   1635
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton CmdExcluirReserva 
      Caption         =   "Excluir"
      Height          =   405
      Left            =   2458
      TabIndex        =   8
      Top             =   4875
      Width           =   1440
   End
   Begin VB.CommandButton CmdCalend2 
      Caption         =   "..."
      Height          =   330
      Left            =   7980
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox TxtDataFinal 
      Height          =   330
      Left            =   6795
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1200
      Width           =   1110
   End
   Begin VB.TextBox TxtTelCel 
      Height          =   330
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1200
      Width           =   1515
   End
   Begin VB.TextBox TxtTelRes 
      Height          =   330
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
   End
   Begin VB.TextBox TxtNomeCliente 
      Height          =   330
      Left            =   5970
      MaxLength       =   30
      TabIndex        =   3
      Top             =   795
      Width           =   2460
   End
   Begin VB.ComboBox CboHorario 
      Height          =   315
      ItemData        =   "Reserva.frx":0000
      Left            =   3525
      List            =   "Reserva.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   795
      Width           =   1245
   End
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "&Gravar Reserva"
      Default         =   -1  'True
      Height          =   405
      Left            =   7028
      TabIndex        =   10
      Top             =   4875
      Width           =   1440
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   405
      Left            =   173
      TabIndex        =   7
      Top             =   4875
      Width           =   1440
   End
   Begin VB.TextBox TxtDataFesta 
      Height          =   330
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   1
      Top             =   795
      Width           =   1140
   End
   Begin VB.CommandButton CmdCalendario 
      Caption         =   "..."
      Height          =   330
      Left            =   2340
      TabIndex        =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Exibir"
      Height          =   195
      Left            =   5580
      TabIndex        =   21
      Top             =   180
      Width           =   375
   End
   Begin VB.Label LblNumReserva 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1950
      TabIndex        =   20
      Top             =   105
      Width           =   2280
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Num. Reserva :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   315
      TabIndex        =   19
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   6015
      TabIndex        =   16
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefone (cel.)"
      Height          =   195
      Left            =   3015
      TabIndex        =   15
      Top             =   1290
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Telefone (res.)"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   1290
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Cliente"
      Height          =   195
      Left            =   4965
      TabIndex        =   13
      Top             =   870
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Horário"
      Height          =   195
      Left            =   2940
      TabIndex        =   12
      Top             =   870
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   735
      TabIndex        =   11
      Top             =   855
      Width           =   345
   End
End
Attribute VB_Name = "Reserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCalend2_Click()

    'TxtDataFinal.Text = RetornaData()
End Sub

Private Sub CmdCalendario_Click()

    'TxtDataFesta.Text = RetornaData()
End Sub

Private Sub CmdExcluirReserva_Click()

    Dim Reserva As New ClsReserva

    'Verificar se foi selecionada uma linha no grid
    If Val(LblNumReserva.Caption) <> 0 Then
        If Reserva.Excluir(Val(LblNumReserva.Caption), Db) Then
            Call LimpaCampos
            Call CarregaGrid
        End If
    End If
End Sub

Private Sub CmdLimpar_Click()

    Call LimpaCampos
    TxtDataFesta.SetFocus
End Sub

Private Sub CmdProcessar_Click()

    Dim Reserva As New ClsReserva

    'Verificar se todos os campos estão preenchidos
    If CamposOK() Then
        If Val(LblNumReserva.Caption) <> 0 Then
            If Reserva.Atualizar(Val(LblNumReserva.Caption), TxtDataFesta.Text, CboHorario.Text, TxtNomeCliente.Text, TxtTelRes.Text, TxtTelCel.Text, TxtDataFinal.Text, Db) Then
                Call LimpaCampos
                Call CarregaGrid
            End If
        Else
            'Inserir Registro de Reserva de Cliente
            If Reserva.Inserir(TxtDataFesta.Text, CboHorario.Text, TxtNomeCliente.Text, TxtTelRes.Text, TxtTelCel.Text, TxtDataFinal.Text, Db) Then
                Call LimpaCampos
                Call CarregaGrid
            End If
        End If
    End If
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Load()

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    'Call CarregaCombo(CboHorario, "HORARIO_FESTA", "ID_HORARIO", "DSC_HORARIO")

    Call FormataGrid
    Call CarregaGrid
End Sub
Private Sub FormataGrid()

    g.Cols = 8

    g.Rows = 2
    g.Row = 0
    
    g.Col = 0
    g.Text = "Num"
    g.ColWidth(0) = 600
    g.ColAlignment(0) = 3

    g.Col = 1
    g.Text = "Data Reserva"
    g.ColWidth(1) = 1100
    g.ColAlignment(1) = 3

    g.Col = 2
    g.Text = "Horário"
    g.ColWidth(2) = 700
    g.ColAlignment(2) = 3

    g.Col = 3
    g.Text = "Nome Cliente"
    g.ColWidth(3) = 2500

    g.Col = 4
    g.Text = "Tel (res)"
    g.ColWidth(4) = 900
    g.ColAlignment(4) = 3

    g.Col = 5
    g.Text = "Tel (cel)"
    g.ColWidth(5) = 900
    g.ColAlignment(5) = 3

    g.Col = 6
    g.Text = "Data Final"
    g.ColWidth(6) = 1000
    g.ColAlignment(6) = 3
    
    g.Col = 7
    g.Text = "OK"
    g.ColWidth(7) = 400
    g.ColAlignment(7) = 3
End Sub
Private Sub CarregaGrid()

    Dim Rs As New ADODB.Recordset

    sSql = "select * from RESERVA order by DATA_RES , HR_RES"
    
    Rs.Open sSql, Db, adOpenDynamic, 1

    g.Rows = 1
    x = 1
    If Not Rs.EOF Then
        Do Until Rs.EOF
            g.Rows = g.Rows + 1
            g.Row = x

            'Num. Reserva
            g.Col = 0
            g.Text = Rs("ID_RES")

            'Data da Festa
            g.Col = 1
            g.Text = Rs("DATA_RES")

            'Horário da Festa
            g.Col = 2
            If Rs("HR_RES") = "A" Then
                g.Text = "Almoço"
            Else
                g.Text = "Noite"
            End If

            'Nome do Cliente interessado
            g.Col = 3
            g.Text = Rs("NOM_CLI")

            'Tel Res
            g.Col = 4
            g.Text = Rs("TEL1_CLI")

            'Tel Cel
            g.Col = 5
            g.Text = Rs("TEL2_CLI")

            'Data Final da Reserva
            g.Col = 6
            g.Text = Rs("DATA_FIM_RES")

            'A reserva virou contrato ?
            g.Col = 7
            g.Text = Rs("FECHOU_CNT") & ""

            x = x + 1
            Rs.MoveNext
        Loop
    End If
End Sub
Private Sub g_Click()

    'Transfere os valores da linha selecionada do grid para os campos de detalhe da tela
    g.Col = 0
    LblNumReserva.Caption = g.Text

    g.Col = 1
    TxtDataFesta.Text = g.Text

    g.Col = 2
    If UCase(g.Text) = "ALMOÇO" Then
        CboHorario.ListIndex = 0
    Else
        CboHorario.ListIndex = 1
    End If

    g.Col = 3
    TxtNomeCliente.Text = g.Text

    g.Col = 4
    TxtTelRes.Text = g.Text

    g.Col = 5
    TxtTelCel.Text = g.Text

    g.Col = 6
    TxtDataFinal.Text = g.Text
End Sub
Private Function CamposOK() As Integer

    CamposOK = True
    'Data da Reserva
    If (Len(Trim(TxtDataFesta.Text)) = 0 Or Not IsDate(TxtDataFesta.Text)) Then
        MsgBox "A Data da Reservada informada não é válida.", vbExclamation, "SGB"
        CamposOK = False
        TxtDataFesta.SetFocus
        Exit Function
    End If

    'Hora da Reserva
    If CboHorario.ListIndex = -1 Then
        MsgBox "É necessário informar o horário a ser Reservado.", vbExclamation, "SGB"
        CamposOK = False
        CboHorario.SetFocus
        Exit Function
    End If

    'Nome do Cliente
    If Len(Trim(TxtNomeCliente.Text)) < 3 Then
        MsgBox "É necessário informar o Nome do Cliente.", vbExclamation, "SGB"
        CamposOK = False
        TxtNomeCliente.SetFocus
        Exit Function
    End If

    'Telefone do cliente
    If Len(Trim(TxtTelRes.Text)) < 8 Then
        MsgBox "É necessário informar o telefone do cliente.", vbExclamation, "SGB"
        CamposOK = False
        TxtTelRes.SetFocus
        Exit Function
    End If
    
    'Data Final da Reserva
    If Len(Trim(TxtDataFinal.Text)) > 0 Then
        If Not IsDate(TxtDataFinal.Text) Then
            MsgBox "A Data Final da Reserva informada não é válida.", vbExclamation, "SGB"
            CamposOK = False
            TxtTelRes.SetFocus
            Exit Function
        End If
    End If
End Function
Private Sub LimpaCampos()

    LblNumReserva.Caption = ""
    TxtDataFesta.Text = ""
    CboHorario.ListIndex = -1
    TxtNomeCliente.Text = ""
    TxtTelRes.Text = ""
    TxtTelCel.Text = ""
    TxtDataFinal.Text = ""
End Sub
