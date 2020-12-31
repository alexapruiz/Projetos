VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "SGB - Sistema de Gerenciamento de Buffets"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu MnuCadClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu MnuCadFornecedores 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu MnuCadColaboradores 
         Caption         =   "Colaboradores"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu MnuReservas 
      Caption         =   "Reservas"
   End
   Begin VB.Menu MnuContrato 
      Caption         =   "Contrato"
   End
   Begin VB.Menu MnuFinanceiro 
      Caption         =   "Financeiro"
      Begin VB.Menu MnuFechamento 
         Caption         =   "&Fechamento"
      End
      Begin VB.Menu MnuContasaPagar 
         Caption         =   "Contas a Pagar"
      End
      Begin VB.Menu MnuSituacao 
         Caption         =   "&Resumo"
      End
      Begin VB.Menu MnuPlanilhaFinanceira 
         Caption         =   "Planilha Financeira"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu MnuConsultaListaContratos 
         Caption         =   "Lista de Contratos"
      End
   End
   Begin VB.Menu MnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu MnuRel1 
         Caption         =   "Contrato"
      End
      Begin VB.Menu MnuRelGerencial 
         Caption         =   "Gerencial"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Financeiro2.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
End Sub
Private Sub MnuCadClientes_Click()

    Cliente.Show 1
End Sub

Private Sub MnuConsultaListaContratos_Click()

    ListaContratos.Show 1
End Sub

Private Sub MnuContasaPagar_Click()

    ContasPagar.Show 1
End Sub
Private Sub MnuContrato_Click()

    Contrato.Show 1
End Sub
Private Sub MnuFechamento_Click()

    Fechamento.Show 1
End Sub
Private Sub MnuPlanilhaFinanceira_Click()

    Financeiro2.Show 1
End Sub
Private Sub MnuRel1_Click()

    Call SelecionaContrato
End Sub

Private Sub MnuRelGerencial_Click()

    DadosGerenciais.Show 1
End Sub

Private Sub MnuReservas_Click()

    Reserva.Show 1
End Sub
Private Sub MnuSair_Click()

    End
End Sub
Private Sub MnuSituacao_Click()

    Financeiro.Show 1
End Sub
Private Sub SelecionaContrato()

    Dim Contrato As String
    Dim sSql As String
    Dim Rec As New ADODB.Recordset

    Contrato = InputBox("Informe o número do contrato para ser impresso", "SGB")

    If Val(Contrato) <> 0 Then
        sSql = "select * from CONTRATOS where id_cnt = " & Contrato

        Rec.Open sSql, Db, adOpenDynamic, adLockOptimistic

        If Not Rec.EOF Then
            Call ImprimeContrato(Rec)
        Else
            MsgBox "Contrato não encontrado", " SGB"
            Exit Sub
        End If
    End If
End Sub
Private Sub ImprimeContrato(ByRef Rec As ADODB.Recordset)

    Dim rec2 As New ADODB.Recordset
    Dim rec3 As New ADODB.Recordset

    'Pesquisa o horário da festa (inicio)
    sSql = "select * from HORARIO_FESTA where id_horario = " & Rec("HR_INI").Value
    rec2.Open sSql, Db, adOpenDynamic, adLockOptimistic
    
    If Not rec2.EOF Then
        'Pesquisa o horário da festa (inicio)
        sSql = "select * from HORARIO_FESTA where id_horario = " & Rec("HR_FIM").Value
        rec3.Open sSql, Db, adOpenDynamic, adLockOptimistic
        If rec3.EOF Then
            MsgBox "Erro ao pesquisar o horário da festa", "SGB"
            Exit Sub
        End If
    Else
        MsgBox "Erro ao pesquisar o horário da festa", "SGB"
        Exit Sub
    End If

    Printer.ScaleMode = 4
    Printer.CurrentX = 5
    Printer.CurrentY = 3
    Printer.FontSize = 20
    Printer.FontName = "Courier New"
    Printer.FontUnderline = True
    Printer.Print ""
    Printer.Print ""
    Printer.FontBold = True
    Printer.Print " BUFFET PLANETA DA ALEGRIA - DADOS DO CONTRATO "
    Printer.FontBold = False
    Printer.Print ""
    Printer.Print ""
    Printer.FontSize = 14
    Printer.Print "Aniversariante  " & Space(11) & " : " & Rec("NOME_ANIV").Value
    Printer.Print ""
    Printer.Print "Data da Festa  " & Space(12) & " : " & Format(Rec("DATA_FESTA").Value, "dd/mm/yyyy")
    Printer.Print ""
    Printer.Print "Horário  " & Space(18) & " : " & rec2("DSC_HORARIO").Value & " às " & rec3("DSC_HORARIO").Value
    Printer.Print ""
    Printer.Print "Idade  " & Space(20) & " : " & Rec("IDADE_ANIV").Value
    Printer.Print ""
    Printer.Print "Qtde Convidados  " & Space(10) & " : " & Rec("QTDE_CONV").Value
    Printer.Print ""
    Printer.Print "Bolo  " & Space(21) & " : " & Rec("DSC_BOLO").Value
    Printer.Print ""
    Printer.Print "Decoração  " & Space(16) & " : " & Rec("DSC_DECOR").Value
    Printer.Print ""
    Printer.Print "Bebida  " & Space(19) & " : " & Rec("OBS_BEBIDA").Value
    Printer.Print ""
    Printer.Print "Nome Pais  " & Space(16) & " : " & Rec("NOM_PAIS").Value
    Printer.Print ""
    Printer.FontSize = 12
    Printer.Print "OBS  " & Space(13) & " : " & Rec("OBS").Value
    Printer.FontSize = 20
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""

    Set rec3 = Nothing

    sSql = "SELECT COL.NOME_COL , FUNC.ID_FUNC , FUNC.DSC_FUNC "
    sSql = sSql & " FROM ESCALA ESC , COLABORADORES COL , FUNCOES FUNC "
    sSql = sSql & " WHERE ID_CNT = " & Rec("ID_CNT").Value
    sSql = sSql & " AND ESC.ID_COL = COL.ID_COL "
    sSql = sSql & " AND ESC.ID_FUNC = FUNC.ID_FUNC "
    sSql = sSql & " ORDER BY FUNC.ID_FUNC "

    rec3.Open sSql, Db, adOpenDynamic, adLockOptimistic
    If rec3.EOF Then
        MsgBox "Erro ao pesquisar a escala", "SGB"
        Exit Sub
    Else
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.Print "Escala de Funcionários"
        Printer.Print ""
        Printer.FontSize = 12
        Printer.Print "Função " & Space(13) & "Nome"
        Printer.FontBold = False
        Do Until rec3.EOF
            Printer.Print Trim(rec3("NOME_COL").Value) & Space(20 - Len(Trim(rec3("NOME_COL").Value))) & rec3("DSC_FUNC").Value
            rec3.MoveNext
        Loop
    End If
    Printer.EndDoc
End Sub
