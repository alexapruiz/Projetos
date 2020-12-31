VERSION 5.00
Begin VB.Form Financeiro2 
   Caption         =   "Financeiro - Planilha"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   2145
   End
   Begin VB.TextBox TxtTotalDespesas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   7320
      TabIndex        =   15
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox TxtTotalReceitas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2970
      TabIndex        =   13
      Top             =   3330
      Width           =   1935
   End
   Begin VB.TextBox TxtResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   5400
      TabIndex        =   11
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   435
      Left            =   5138
      TabIndex        =   10
      Top             =   5280
      Width           =   1725
   End
   Begin VB.TextBox TxtCustosFixos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2340
      Width           =   2145
   End
   Begin VB.TextBox TxtDespesasFestas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1830
      Width           =   2145
   End
   Begin VB.TextBox TxtQtdeFestas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1350
      Width           =   2145
   End
   Begin VB.TextBox TxtReceitas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox CboMes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5970
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   2385
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Outras Despesas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5610
      TabIndex        =   18
      Top             =   2850
      Width           =   1575
   End
   Begin VB.Line Line7 
      X1              =   1200
      X2              =   10440
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total Despesas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5520
      TabIndex        =   16
      Top             =   3420
      Width           =   1665
   End
   Begin VB.Line Line5 
      X1              =   1200
      X2              =   1200
      Y1              =   3960
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   10440
      X2              =   10440
      Y1              =   1080
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   10440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   5040
      Y1              =   3960
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   10440
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Receitas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   3390
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4080
      TabIndex        =   12
      Top             =   4500
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Custos Fixos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6045
      TabIndex        =   5
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Despesas (estimado)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5235
      TabIndex        =   4
      Top             =   1890
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qtde Festas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6090
      TabIndex        =   3
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Receitas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mês de Referência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      TabIndex        =   0
      Top             =   510
      Width           =   2385
   End
End
Attribute VB_Name = "Financeiro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Db2 As New ADODB.Connection
Private Sub Command1_Click()

    MsgBox CboMes.ListIndex + 1
End Sub
Private Sub CboMes_Click()

    Call CalculaValores
End Sub
Private Sub CmdSair_Click()

    Unload Me
End Sub
Private Sub Form_Load()

    Db2.Open "Financas"

    Call PreencheComboMes
End Sub
Private Sub PreencheComboMes()

    CboMes.AddItem "Janeiro"
    CboMes.AddItem "Fevereiro"
    CboMes.AddItem "Março"
    CboMes.AddItem "Abril"
    CboMes.AddItem "Maio"
    CboMes.AddItem "Junho"
    CboMes.AddItem "Julho"
    CboMes.AddItem "Agosto"
    CboMes.AddItem "Setembro"
    CboMes.AddItem "Outubro"
    CboMes.AddItem "Novembro"
    CboMes.AddItem "Dezembro"
    CboMes.AddItem "Janeiro/2007"
    CboMes.AddItem "Fevereiro/2007"
    CboMes.AddItem "Março/2007"
    CboMes.AddItem "Abril/2007"
    
End Sub
Private Sub CalculaValores()

    Dim Rs      As New ADODB.Recordset
    Dim Rs2     As New ADODB.Recordset
    Dim sSql    As String
    Dim TABELA  As String
    Dim NumDias As String
    Dim Mes     As String
    Dim VALOR   As Currency

    'Selecionar qual tabela ler
    Select Case CboMes.ListIndex + 1
        Case 1
            TABELA = "01_Janeiro"
        Case 2
            TABELA = "02_Fevereiro"
        Case 3
            TABELA = "03_Marco"
        Case 4
            TABELA = "04_Abril"
            NumDias = 30
        Case 5
            TABELA = "05_Maio"
            NumDias = 31
        Case 6
            TABELA = "06_Junho"
            NumDias = 30
        Case 7
            TABELA = "07_Julho"
            NumDias = 31
        Case 8
            TABELA = "08_Agosto"
            NumDias = 31
        Case 9
            TABELA = "09_Setembro"
            NumDias = 30
        Case 10
            TABELA = "10_Outubro"
            NumDias = 31
        Case 11
            TABELA = "11_Novembro"
            NumDias = 30
        Case 12
            TABELA = "12_Dezembro"
            NumDias = 31
        Case 13
            TABELA = "13_Janeiro"
            NumDias = 31
        Case 14
            TABELA = "14_Fevereiro"
            NumDias = 29
        Case 15
            TABELA = "15_Marco"
            NumDias = 31
        Case 16
            TABELA = "16_Abril"
            NumDias = 30
    End Select

    'Calcular o valor total das receitas ******************************************************************
    sSql = "select sum(valor) as Total from " & TABELA

    Rs.Open sSql, Db2, adOpenStatic, adLockOptimistic

    If Not IsNull(Rs("Total").Value) Then
        TxtReceitas.Text = Rs("Total").Value
    Else
        MsgBox "Erro ao calcular o Valor Total das Receitas", vbOKOnly, "SGB"
        Exit Sub
    End If
    Set Rs = Nothing
    '******************************************************************************************************

    'Calcular a quantidade de festas do mes ***************************************************************
    Mes = Format(CboMes.ListIndex + 1, "00")

    sSql = "select count(0) as Qtde_Festas from Festas_2006 "
    sSql = sSql & " where data between #" & Mes & "/01/2006# and #" & Mes & "/" & NumDias & "/2006#"

    Rs.Open sSql, Db2, adOpenStatic, adLockOptimistic

    TxtQtdeFestas.Text = Rs("Qtde_Festas").Value
    Set Rs = Nothing
    '******************************************************************************************************

    'Custo Fixo *******************************************************************************************
    TxtCustosFixos.Text = "6500,00"
    '******************************************************************************************************

    'Calcular o valor estimado para as festas do mes selecionado ******************************************
    sSql = "select * from Festas_2006 "
    sSql = sSql & " where data between #" & Mes & "/01/2006# and #" & Mes & "/" & NumDias & "/2006#"

    Rs.Open sSql, Db2, adOpenStatic, adLockOptimistic
    Do Until Rs.EOF
        sSql = "select valor from preco_custo "
        sSql = sSql & " where tipo = '" & Rs("tipo").Value & "' and qtde = " & Rs("qtde").Value
        
        Rs2.Open sSql, Db2, adOpenStatic, adLockOptimistic

        If Not Rs2.EOF Then
            VALOR = VALOR + Rs2("valor").Value
        Else
            MsgBox "Erro ao pesquisar valor do preço de custo da festa", vbOKOnly, "SGB"
            Exit Sub
        End If
        Set Rs2 = Nothing

        Rs.MoveNext
    Loop
    TxtDespesasFestas.Text = Format(VALOR, ".00")
    '******************************************************************************************************

    'Calcular os totais ***********************************************************************************
    TxtTotalReceitas.Text = TxtReceitas.Text
    TxtTotalDespesas.Text = Format(CCur(TxtCustosFixos) + CCur(TxtDespesasFestas.Text), ".00")
    '******************************************************************************************************
    
    'Apurar se houve LUCRO ou PREJUÍZO
    TxtResultado.Text = CCur(TxtTotalReceitas.Text) - CCur(TxtTotalDespesas.Text)
    If CCur(TxtResultado.Text) > 0 Then
        'Lucro
        TxtResultado.ForeColor = RGB(0, 0, 255)
    Else
        'Prejuízo
        TxtResultado.ForeColor = RGB(255, 0, 0)
    End If
End Sub

