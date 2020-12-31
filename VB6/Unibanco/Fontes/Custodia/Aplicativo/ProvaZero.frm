VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ProvaZero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prova Zero"
   ClientHeight    =   8010
   ClientLeft      =   810
   ClientTop       =   600
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10395
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   72
      Top             =   6720
   End
   Begin VB.Frame fraCheques 
      Caption         =   "Cheques por Depósito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3444
      Left            =   432
      TabIndex        =   8
      Top             =   4152
      Width           =   9660
      Begin MSFlexGridLib.MSFlexGrid grdCheques 
         Height          =   3030
         Left            =   285
         TabIndex        =   9
         Top             =   240
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   8388608
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame fraSuperior 
      Height          =   3972
      Left            =   408
      TabIndex        =   7
      Top             =   72
      Width           =   9660
      Begin VB.Frame fraInformacaoDeposito 
         Caption         =   "Total do Borderô"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1356
         Left            =   384
         TabIndex        =   13
         Top             =   2448
         Width           =   6612
         Begin VB.Label lblVlrChequesCalculado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   4608
            TabIndex        =   21
            Top             =   912
            Width           =   1812
         End
         Begin VB.Label lblQtdChequesCalculado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   4608
            TabIndex        =   20
            Top             =   552
            Width           =   1812
         End
         Begin VB.Label lblCalcudado 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Calculado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   4608
            TabIndex        =   19
            Top             =   168
            Width           =   1812
         End
         Begin VB.Label lblInformado 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Informado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   2496
            TabIndex        =   18
            Top             =   168
            Width           =   1812
         End
         Begin VB.Label lblVlrChequesInformado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   2496
            TabIndex        =   17
            Top             =   912
            Width           =   1812
         End
         Begin VB.Label lblQtdChequesInformado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   2496
            TabIndex        =   16
            Top             =   552
            Width           =   1812
         End
         Begin VB.Label lblVlrDeposito 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor do Depósito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   384
            TabIndex        =   15
            Top             =   912
            Width           =   1812
         End
         Begin VB.Label lblQtdCheque 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Qtd. Cheques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   384
            TabIndex        =   14
            Top             =   576
            Width           =   1812
         End
      End
      Begin VB.Frame fraDataDeposito 
         Caption         =   "Diferença em Depósito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2196
         Left            =   2688
         TabIndex        =   11
         Top             =   192
         Width           =   4308
         Begin MSFlexGridLib.MSFlexGrid grdDataDeposito 
            Height          =   1884
            Left            =   192
            TabIndex        =   12
            Top             =   216
            Width           =   3924
            _ExtentX        =   6906
            _ExtentY        =   3334
            _Version        =   393216
            Rows            =   10
            FixedCols       =   0
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
      Begin VB.Frame fraBordero 
         Caption         =   "Borderôs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2196
         Left            =   384
         TabIndex        =   10
         Top             =   192
         Width           =   2148
         Begin MSFlexGridLib.MSFlexGrid grdBordero 
            Height          =   1884
            Left            =   120
            TabIndex        =   22
            Top             =   216
            Width           =   1908
            _ExtentX        =   3360
            _ExtentY        =   3334
            _Version        =   393216
            Rows            =   10
            Cols            =   1
            FixedCols       =   0
            BackColorSel    =   8388608
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
      Begin VB.Frame fraBotoes 
         Height          =   3636
         Left            =   7296
         TabIndex        =   6
         Top             =   168
         Width           =   2124
         Begin VB.CommandButton cmdAtualizaTela 
            Caption         =   "&Atualizar Tela"
            Height          =   396
            Left            =   288
            TabIndex        =   4
            Top             =   1848
            Width           =   1572
         End
         Begin VB.CommandButton cmdIncluirCheque 
            Caption         =   "&Incluir Cheque"
            Height          =   396
            Left            =   288
            TabIndex        =   1
            Top             =   336
            Width           =   1572
         End
         Begin VB.CommandButton cmdExcluirCheque 
            Caption         =   "&Excluir Cheque"
            Height          =   396
            Left            =   288
            TabIndex        =   2
            Top             =   840
            Width           =   1572
         End
         Begin VB.CommandButton cmdEnviaSupervisor 
            Caption         =   "Enviar p/ Su&pervisor"
            Height          =   396
            Left            =   288
            TabIndex        =   3
            Top             =   1344
            Width           =   1572
         End
         Begin VB.CommandButton cmdEncerrarCapa 
            Caption         =   "En&cerrar Capa"
            Height          =   396
            Left            =   288
            TabIndex        =   5
            Top             =   2376
            Width           =   1572
         End
         Begin VB.CommandButton cmdSair 
            Caption         =   "&Sair"
            Height          =   396
            Left            =   312
            TabIndex        =   0
            Top             =   3024
            Width           =   1572
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   228
      Left            =   7608
      TabIndex        =   23
      Top             =   7752
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   24
      Top             =   7716
      Width           =   10392
      _ExtentX        =   18336
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13229
            MinWidth        =   13229
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5027
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ProvaZero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Proc_Selecionar      As New Custodia.Selecionar
Dim Proc_atualizar       As New Custodia.Atualizar
Dim rsBordero            As New ADODB.Recordset
Dim RsDeposito           As New ADODB.Recordset
Dim RsCheque             As New ADODB.Recordset
Dim DataProcessamento    As String
Dim objAguardaDocumento  As AguardaDocumento

Private sTempo           As Integer               'Controle de tempo para ativar (Timer) informando Borderô em Prova Zero

Private Coluna As TpColunas
Private Type TpColunas
     'Número de Colunas (Grid Borderôs)
     Bor_NrBord          As Integer
     Bor_QtdCheques      As Integer
     Bor_VlrCheques      As Integer
     Bor_Supervisor      As Integer
     
     'Número de Colunas (Grid Depósitos)
     Dep_Data            As Integer
     Dep_QtdDifer        As Integer
     Dep_VlrDifer        As Integer
     Dep_QtdDeposito     As Integer
     Dep_VlrDeposito     As Integer
     Dep_QtdTabCheque    As Integer
     Dep_VlrTabCheque      As Integer
     
     'Número de Colunas (Grid Cheques)
     Chq_NrSeq           As Integer
     Chq_NrBanco         As Integer
     Chq_NrAgencia       As Integer
     Chq_NrConta         As Integer
     Chq_NrCheque        As Integer
     Chq_CPF_CNPJ        As Integer
     Chq_VlrCheque       As Integer
     
End Type

Private Acumulador As tpAcumulador
Private Type tpAcumulador
     lQtdTotBordero      As Long
     dVlrTotBordero      As Double
     lQtdTotCheques      As Long
     dVlrTotCheques      As Double
End Type

Private Ambiente As tpAmbiente
Private Type tpAmbiente
     iBor_RowSel         As Integer          'Contém o número da linha do grid de borderô selecionado (Marcação)
     bBor_Updated        As Boolean          'Informa se houve manutenção nas informações do borderô selecionado
     iDep_RowSel         As Integer          'Contém o número da linha do grid da Data Depósito selecionado (Marcação)
     iChq_RowSel         As Integer          'Contém o número da linha do grid de Cheque selecionado (Marcação)
End Type

Private Sub cmdAtualizaTela_Click()

     If Ambiente.iBor_RowSel <> 0 Then
          Call VoltaStatusBordero(grdBordero, Ambiente.iBor_RowSel)
     End If

     'Carrega grid com os borderôs com status = (4)Para Prova Zero
     Call LoadGridBordero

End Sub

Private Sub cmdEncerrarCapa_Click()

Dim i          As Integer
Dim lRetorno   As Long
Dim Fechamento As New CalculoBordero       'Classe de calculo (fechamento)

On Error GoTo Erro_cmdEncerrarCapa_Click

     'Verifica totais do borderô em relação ao totais calculado
     If Str(Acumulador.dVlrTotBordero) <> Str(Acumulador.dVlrTotCheques) Then
          Beep
          MsgBox "Valores divergentes na somatória geral.", vbInformation, Me.Caption
          Exit Sub
     End If
     
     If Acumulador.lQtdTotBordero <> Acumulador.lQtdTotCheques Then
          Beep
          MsgBox "Quantidade divergente na somatória geral.", vbInformation, Me.Caption
          Exit Sub
     End If

     'Verifica se totais por data depósito zera com totais de cheque por data depósito
     With grdDataDeposito
          Call SelecionaLinha(grdDataDeposito)
          
          For i = 1 To .Rows - 1
               .Row = i
               .Col = Coluna.Dep_QtdDifer
               If .Text < 0 Then Exit For
               .Col = Coluna.Dep_VlrDifer
               If .Text < 0 Then Exit For
          Next
          
          'Verifica se existe divergencia na quantidade
          If .Col = Coluna.Dep_QtdDifer And .Text < 0 Then
               grdDataDeposito_Click
               Beep
               MsgBox "Quantidade divergente na somatória por data depósito.", vbInformation, Me.Caption
               Exit Sub
          End If
          
          'Verifica se existe divergencia no valor
          If .Col = Coluna.Dep_VlrDifer And .Text < 0 Then
               grdDataDeposito_Click
               Beep
               MsgBox "Valores divergentes na somatória por data depósito.", vbInformation, Me.Caption
               Exit Sub
          End If
          
          .Row = 1
          Call SelecionaLinha(grdDataDeposito, True)
     End With

     
     'Inicializa transação
     g_cMainConnection.BeginTrans
     Screen.MousePointer = vbHourglass
     
     'Processa verificação de cheques/datas indevidas
     Fechamento.SetConnection g_cMainConnection
     Fechamento.IdBordero = grdBordero.RowData(grdBordero.Row)
     Fechamento.DataProcessamento = Geral.DataProcessamento
     
     Fechamento.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
     Fechamento.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas

     Call Fechamento.VoltaStatusChequesIndevidos

     Call Fechamento.CalculaChequesIndevidosQTDE
     Call Fechamento.CalculaChequesIndevidosDATA
     
     'Altera Status do borderô de (G)Em Prova Zero para (R)Transmissâo
     Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), "R", "G"), lRetorno, adCmdText)

     'Verifica se conseguiu enviar o borderô para supervisor
     If lRetorno = 0 Then
          'Finaliza transação
          g_cMainConnection.RollbackTrans
          Screen.MousePointer = vbDefault
          
          Beep
          MsgBox "Não foi possível encerrar este borderô, favor atualizar a tela e tentar novamente.", vbCritical, Me.Caption
          Exit Sub
     End If
     
     'Encerra transação
     g_cMainConnection.CommitTrans
     
     'Finaliza controle de tempo para informar Borderô Em Prova Zero
     sTempo = 0
     tmrAtualiza.Enabled = False
     
     If grdBordero.Rows > 2 Then
          grdBordero.RemoveItem (grdBordero.Row)
     Else
          grdBordero.Rows = 1
     End If
     
     'Limpa grid´s de Borderô, Depósito e Cheques
     Call LimpaTela
     
     'Seta cursor para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_cmdEncerrarCapa_Click:

     'Encerra transação
     g_cMainConnection.RollbackTrans
     
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível encerrar este borderô, favor atualizar a tela e tentar novamente.", vbCritical, Me.Caption
     

End Sub

Private Sub cmdEnviaSupervisor_Click()

Dim lRetorno As Long

On Error GoTo Erro_cmdEnviaSupervisor

     With grdBordero
          .Col = Coluna.Bor_Supervisor
          If .Text Then
               .Col = Coluna.Bor_NrBord
               Beep
               MsgBox "Borderô já enviado para supervisor", vbInformation, Me.Caption
               Exit Sub
          End If
     End With
     
     Call SelecionaLinha(grdBordero, True)
     
     'Inicializa transação
     g_cMainConnection.BeginTrans
     Screen.MousePointer = vbHourglass
     
     
     'Altera Status do borderô de (G)Em Prova Zero para (5)Supervisor
     Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), "5", "G"), lRetorno, adCmdText)

     'Verifica se conseguiu enviar o borderô para supervisor
     If lRetorno = 0 Then
          'Finaliza transação
          g_cMainConnection.RollbackTrans
          Screen.MousePointer = vbDefault
          
          Beep
          MsgBox "Não foi possível enviar este borderô para supervisor, favor atualizar a tela", vbCritical, Me.Caption
          Exit Sub
     End If
     
     'Encerra transação
     g_cMainConnection.CommitTrans
     
     'Finaliza controle de tempo para informar Borderô Em Prova Zero
     sTempo = 0
     tmrAtualiza.Enabled = False
     
     If grdBordero.Rows > 2 Then
          grdBordero.RemoveItem (grdBordero.Row)
     Else
          grdBordero.Rows = 1
     End If
     
     'Limpa grid´s de Borderô, Depósito e Cheques
     Call LimpaTela
     
     'Seta cursor para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_cmdEnviaSupervisor:
     
     'Encerra transação
     g_cMainConnection.RollbackTrans
     
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível enviar o borderô para supervisor", vbCritical, Me.Caption

End Sub

Private Sub cmdExcluirCheque_Click()

Dim lRetorno        As Long
Dim dVlrDeposito    As Double
Dim lQtdDeposito    As Long
Dim dVlrCheques     As Double
Dim lQtdCheques     As Long
Dim dValorDoCheque  As Double
Dim sDelBco         As String, sDelAge As String, sDelConta As String, sDelChq As String

On Error GoTo Erro_cmdExcluirCheque_Click

     'Inicializa variáveis
     dVlrDeposito = 0
     lQtdDeposito = 0
     dVlrCheques = 0
     lQtdCheques = 0

     grdCheques.Col = Coluna.Chq_NrBanco:    sDelBco = Trim(grdCheques.Text)
     grdCheques.Col = Coluna.Chq_NrAgencia:  sDelAge = Trim(grdCheques.Text)
     grdCheques.Col = Coluna.Chq_NrConta:    sDelConta = Trim(grdCheques.Text)
     grdCheques.Col = Coluna.Chq_VlrCheque:  sDelChq = Trim(grdCheques.Text)
          
     If MsgBox("Confirma exclusão do cheque," & vbCrLf & vbCrLf & _
          "Banco:      " & sDelBco & vbCrLf & vbCrLf & _
          "Agência:    " & sDelAge & vbCrLf & vbCrLf & _
          "Conta:      " & sDelConta & vbCrLf & vbCrLf & _
          "Valor:      " & sDelChq & vbCrLf & vbCrLf _
          , vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
          Exit Sub
     End If
     
     'Inicializa transação
     g_cMainConnection.BeginTrans
     Screen.MousePointer = vbHourglass
     
     'Altera Status do borderô de (G)Em Prova Zero para (5)Supervisor
     Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusCheque(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), grdCheques.RowData(grdCheques.Row), "D"), lRetorno, adCmdText)

     'Verifica se conseguiu enviar o borderô para supervisor
     If lRetorno = 0 Then
          'Finaliza transação
          g_cMainConnection.RollbackTrans
          Screen.MousePointer = vbDefault
          
          Beep
          MsgBox "Não foi possível excluir este cheque. Verificar!", vbCritical, Me.Caption
          Exit Sub
     End If
     
     'Encerra transação
     g_cMainConnection.CommitTrans
     
     'Obtem o valor do cheque
     grdCheques.Col = Coluna.Chq_VlrCheque
     dValorDoCheque = grdCheques.Text
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '    Atualiza acumuladores para calculo de totais do borderô     '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     With grdDataDeposito
          'Obtem Qtd e Vlr. de depósito
          .Col = Coluna.Dep_QtdDeposito:     lQtdDeposito = .Text
          .Col = Coluna.Dep_VlrDeposito:     dVlrDeposito = .Text
          
          'Obtem o total (QTD e VLR) de cheques na data de depósito
          .Col = Coluna.Dep_QtdTabCheque:    lQtdCheques = .Text
          .Col = Coluna.Dep_VlrTabCheque:    dVlrCheques = .Text
          
          lQtdCheques = lQtdCheques - 1
          dVlrCheques = dVlrCheques - dValorDoCheque
          
          'Acerta coluna com somatória de cheques
          .Col = Coluna.Dep_QtdTabCheque:    .Text = lQtdCheques
          .Col = Coluna.Dep_VlrTabCheque:    .Text = dVlrCheques
          
          'Acumula total de Qtd. e Valor para apresentar em Total do Borderô
          Acumulador.lQtdTotCheques = Acumulador.lQtdTotCheques - 1
          Acumulador.dVlrTotCheques = Acumulador.dVlrTotCheques - dValorDoCheque
          
          'Calcula diferença entre valor de depósito e valor total de cheques
          .Col = Coluna.Dep_QtdDifer:   .Text = Formato(lQtdCheques - lQtdDeposito, "I")
          .Col = Coluna.Dep_VlrDifer:   .Text = Formato(dVlrCheques - dVlrDeposito)
          
          
          'Colorir a coluna do grid conforme Qtd. e Valor
          Call ColorirColuna(grdDataDeposito, Coluna.Dep_QtdDifer)
          Call ColorirColuna(grdDataDeposito, Coluna.Dep_VlrDifer)
     End With
     
     If grdCheques.Rows > 2 Then
          grdCheques.RemoveItem (grdCheques.Row)
     Else
          grdCheques.Rows = 1
          cmdExcluirCheque.Enabled = False
     End If
     
     Call SelecionaLinha(grdDataDeposito, True)
     Call SelecionaLinha(grdCheques, True)
     Call LoadGridTotalBordero

     'Seta cursor para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_cmdExcluirCheque_Click:
     
     'Encerra transação
     g_cMainConnection.RollbackTrans
     
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível excluir este cheque. Verificar!", vbInformation, Me.Caption

End Sub

Private Sub cmdIncluirCheque_Click()

Dim IdCheque As Double, IdBordero As Long, iCancel As enumRetornoModal
Dim sDataDeposito As String, i As Integer, bMesmaDataDep As Boolean, bMesmoCheque As Boolean
Dim lngDataDeposito As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VERIFICAR PARAMETRO COM TOTAL DE CHEQUES POR BORDERO
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (Acumulador.lQtdTotCheques + 1) > g_Parametros.QuantidadeCheques Then
     MsgBox "Inclusão de cheque não permitido." & vbCrLf & vbCrLf & _
     "Limite máximo de cheques por borderô = " & g_Parametros.QuantidadeCheques, vbInformation, Me.Caption
     Exit Sub
End If

'Obtem o IdBordero do Bordero selecionado
IdBordero = grdBordero.RowData(grdBordero.Row)

'Inicializa variavel de retorno com IdCheque do novo cheque incluido
IdCheque = 0

'Obtem a Data de depósito do grid Depósito
grdDataDeposito.Col = Coluna.Dep_Data
lngDataDeposito = FormataAMD(grdDataDeposito.Text)
sDataDeposito = grdDataDeposito.Text

'Habilita seleção por linha
Call SelecionaLinha(grdDataDeposito, True)

'Setar parâmetros do form (CHEQUE) para inclusão de cheque
Cheque.SetIdCheque (0)
iCancel = Cheque.ShowModal(IdBordero, lngDataDeposito, IdCheque)

'Se Encerrou com modificação, altera grid´s de Data e Cheque
If iCancel = eRetornoOK Then
     'Muda ponteiro do windows
     Screen.MousePointer = vbHourglass
     
     grdCheques.Rows = 1
     Call SelecionaLinha(grdCheques)
     
     'Carrega grid de Depósito do borderô
     Call LoadGridDeposito

     bMesmaDataDep = False
     For i = 1 To grdDataDeposito.Rows - 1
          grdDataDeposito.Row = i
          If grdDataDeposito.Text = sDataDeposito Then
               bMesmaDataDep = True
               Exit For
          End If
     Next
     
     'Posiciona na mesma Data de Depósito do grid (Data Deposito)
     If bMesmaDataDep Then
          'Habilita seleção por linha
          grdDataDeposito.Row = i
          Call SelecionaLinha(grdDataDeposito, True)
          'Carrega grid com cheques referentes a mesma data de depósito
          Call LoadGridCheques
     
     Else
          If grdDataDeposito.Rows = 1 Then
               grdDataDeposito.Row = 1
               'Desabilita seleção por linha
               Call SelecionaLinha(grdDataDeposito)
          Else
               grdDataDeposito.Row = 2
               'Habilita seleção por linha
               Call SelecionaLinha(grdDataDeposito, True)
          End If
     End If
     
     'Apresenta frame com totais do Borderô
     Call LoadGridTotalBordero

     'Posiciona no mesmo registro do cheque do grid (CHEQUES)
     bMesmoCheque = False
     For i = 1 To grdCheques.Rows - 1
          If grdCheques.RowData(i) = IdCheque Then
               grdCheques.Row = i
               bMesmoCheque = True
               Exit For
          End If
     Next
     
     If bMesmoCheque Then
          'Habilita seleção por linha
          Call SelecionaLinha(grdCheques, True)
          Exit Sub
     Else
          If grdCheques.Rows = 1 Then
               grdCheques.Row = 0
               'Desabilita seleção por linha
               Call SelecionaLinha(grdCheques)
          Else
               grdCheques.Row = 2
               'Habilita seleção por linha
               Call SelecionaLinha(grdCheques, True)
          End If
     End If
End If

'Muda ponteiro do windows para default
Screen.MousePointer = vbDefault

End Sub

Private Sub Sair()

Dim lRetorno As Long

On Error GoTo Erro_Sair

     'Verifica se borderô bloqueado para em prova zero(Modo Seleção) ou se existe borderô no grid
     If grdBordero.HighLight = flexHighlightNever Or grdBordero.Rows <= 1 Then
          If Not objAguardaDocumento Is Nothing Then
              objAguardaDocumento.Finalizar
          End If
          
          Unload Me
          Exit Sub
     End If

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '         Atualiza status do Borderô para  Prova Zero            '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     g_cMainConnection.BeginTrans
     
     Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), "4", "G"), lRetorno, adCmdText)
     
     g_cMainConnection.CommitTrans

     'Finaliza controle de tempo para informar Borderô Em Prova Zero
     sTempo = 0
     tmrAtualiza.Enabled = False

     If Not objAguardaDocumento Is Nothing Then
         objAguardaDocumento.Finalizar
     End If
     
     Unload Me
     Exit Sub
     
Erro_Sair:

     'Rollback na transação
     g_cMainConnection.RollbackTrans
     
     Beep
     MsgBox Err.Description, vbCritical, Me.Caption

End Sub

Private Sub cmdSair_Click()
     
     Unload Me
     
End Sub

Private Sub Form_Activate()

'Verifica se houve chamada do form Borderô/Cheque então não carregar grid borderô
If grdBordero.RowData(grdBordero.Row) = 0 Then
     'Carrega grid com os borderôs com status = (4)Para Prova Zero
     Call LoadGridBordero
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then Call Sair

End Sub

Private Sub Form_Load()

Me.Caption = App.Title & "  -  " & Me.Caption

'Número de Colunas (Grid Borderôs)
Coluna.Bor_NrBord = 0
Coluna.Bor_QtdCheques = 1
Coluna.Bor_VlrCheques = 2
Coluna.Bor_Supervisor = 3

'Número de Colunas (Grid Depósitos)
Coluna.Dep_Data = 0
Coluna.Dep_QtdDifer = 1
Coluna.Dep_VlrDifer = 2
Coluna.Dep_QtdDeposito = 3
Coluna.Dep_VlrDeposito = 4
Coluna.Dep_QtdTabCheque = 5
Coluna.Dep_VlrTabCheque = 6

'Número de Colunas (Grid Cheques)
Coluna.Chq_NrSeq = 0
Coluna.Chq_NrBanco = 1
Coluna.Chq_NrAgencia = 2
Coluna.Chq_NrConta = 3
Coluna.Chq_NrCheque = 4
Coluna.Chq_CPF_CNPJ = 5
Coluna.Chq_VlrCheque = 6

''''''''''''''''''''''''''''''''''''
'         Grid de Borderôs         '
''''''''''''''''''''''''''''''''''''
With grdBordero
     .Rows = 1
     .Cols = 4
     .ColWidth(0) = 1850
     'Esconde coluna com IdBordero
     .Row = 0
     .Col = Coluna.Bor_NrBord:     .Text = "Número"
     .Col = Coluna.Bor_QtdCheques: .Text = "Qtd Cheques"
     .Col = Coluna.Bor_VlrCheques: .Text = "Vlr Cheques"
     .ColAlignment(Coluna.Bor_NrBord) = flexAlignCenterCenter
     .BackColorSel = vbBlack
     .SelectionMode = flexSelectionByRow
     .ScrollBars = flexScrollBarVertical
     Call SelecionaLinha(grdBordero)
End With

'''''''''''''''''''''''''''''''''''''''''
'         Grid Data de Depósito         '
'''''''''''''''''''''''''''''''''''''''''
With grdDataDeposito
     .Rows = 1: .Cols = 7
     .Row = 0
     .ColWidth(Coluna.Dep_Data) = 1100
     .ColWidth(Coluna.Dep_QtdDifer) = 1100
     .ColWidth(Coluna.Dep_VlrDifer) = 1630

     .Col = Coluna.Dep_Data:            .Text = "Data"
     .Col = Coluna.Dep_QtdDifer:        .Text = "Quantidade"
     .Col = Coluna.Dep_VlrDifer:        .Text = "Valor"
     .Col = Coluna.Dep_QtdDeposito:     .Text = "Qtd Deposito"
     .Col = Coluna.Dep_VlrDeposito:     .Text = "Vlr Deposito"
     .Col = Coluna.Dep_QtdTabCheque:    .Text = "Qtd Tab Chq"
     .Col = Coluna.Dep_VlrTabCheque:    .Text = "Vlr Tab Chq"

     .ColAlignment(Coluna.Dep_Data) = flexAlignCenterCenter
     .ColAlignment(Coluna.Dep_QtdDifer) = flexAlignCenterCenter
     .ColAlignment(Coluna.Dep_VlrDifer) = flexAlignRightCenter

     .BackColorSel = vbBlack
     .SelectionMode = flexSelectionByRow
     .ScrollBars = flexScrollBarVertical
     Call SelecionaLinha(grdDataDeposito)
End With

''''''''''''''''''''''''''''''''''''
'         Grid de Cheques          '
''''''''''''''''''''''''''''''''''''
With grdCheques
     .Rows = 1: .Cols = 7
     .Row = 0
     .ColWidth(Coluna.Chq_NrSeq) = 800
     .ColWidth(Coluna.Chq_NrBanco) = 700
     .ColWidth(Coluna.Chq_NrAgencia) = 1300
     .ColWidth(Coluna.Chq_NrConta) = 1200
     .ColWidth(Coluna.Chq_NrCheque) = 1400
     .ColWidth(Coluna.Chq_CPF_CNPJ) = 2000        'Se alterar largura desta coluna, acertar tbem na sub (LoadGridCheque)
     .ColWidth(Coluna.Chq_VlrCheque) = 1690

     .Col = Coluna.Chq_NrSeq:      .Text = "Nr."
     .Col = Coluna.Chq_NrBanco:    .Text = "Banco"
     .Col = Coluna.Chq_NrAgencia:  .Text = "Agência"
     .Col = Coluna.Chq_NrConta:    .Text = "Conta"
     .Col = Coluna.Chq_NrCheque:   .Text = "Nr. Cheque"
     .Col = Coluna.Chq_CPF_CNPJ:   .Text = "CPF/CNPJ"
     .Col = Coluna.Chq_VlrCheque:  .Text = "Valor"

     .ColAlignment(Coluna.Chq_NrSeq) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_NrBanco) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_NrAgencia) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_NrConta) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_NrCheque) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_CPF_CNPJ) = flexAlignCenterCenter
     .ColAlignment(Coluna.Chq_VlrCheque) = flexAlignRightCenter

     .BackColorSel = vbBlack
     .SelectionMode = flexSelectionByRow
     Call SelecionaLinha(grdCheques)

End With

'Acerta botôes
cmdIncluirCheque.Enabled = False
cmdExcluirCheque.Enabled = False
cmdEnviaSupervisor.Enabled = False
cmdAtualizaTela.Enabled = True
cmdEncerrarCapa.Enabled = False

'Converte data de processamento para (dd/mm/aaaa)
DataProcessamento = FormataDMA(Geral.DataProcessamento)

'Centraliza o form
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

'Inicializa scanner
Call Principal.SetScanner

End Sub

Private Function FormataDMA(ByVal lngData As Long) As String

     'Converte para data (dd/mm/yyyy)
     FormataDMA = Format(Right(lngData, 2) & "/" & Mid(lngData, 5, 2) & "/" & Left(lngData, 4), "dd/mm/yyyy")

End Function

Private Function FormataAMD(ByVal strData As String) As Long
     
     'Converte para data (yyyymmdd)
     FormataAMD = Right(strData, 4) & Mid(strData, 4, 2) & Left(strData, 2)

End Function
Private Function Formato(ByVal Valor As Variant, Optional ByVal TipoFormato As String = "V") As String
'TipoFormato = (V)alor / (I)inteiro

If TipoFormato = "V" Then
     Formato = Format(Valor, "###,###,##0.00")
Else
     Formato = Format(Valor, "##,###,##0")
End If

End Function

Private Sub ColorirColuna(objeto As Object, ByVal Col As Integer)

Dim iColAnt    As Integer

'Guarda posição anterior da coluna
iColAnt = objeto.Col

objeto.Col = Col

If objeto.Text <> 0 Then
     objeto.CellForeColor = vbRed
End If
     
Sair:
     objeto.Col = iColAnt

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Finaliza Scanner
Call Principal.DelScanner
 
Call Sair
 
End Sub

Private Sub grdBordero_Click()

Dim lRetorno As Long

On Error GoTo Erro_grdBordero_Click

     'Verifica se seleção do borderô já efetuada
     If Ambiente.iBor_RowSel = grdBordero.Row Then Exit Sub
     
     If Ambiente.iBor_RowSel <> 0 Then
          Call VoltaStatusBordero(grdBordero, Ambiente.iBor_RowSel)
     End If
     
     'Seta ponteiro para aguardo..
     Screen.MousePointer = vbHourglass
     
     'Habilita seleção por linha
     Call SelecionaLinha(grdBordero, True)

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '         Atualiza status do Borderô para em Prova Zero          '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     g_cMainConnection.BeginTrans
     
     Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), "G", "4"), lRetorno, adCmdText)
     
     g_cMainConnection.CommitTrans

     If lRetorno = 0 Then
          Call LimpaTela
          Call SelecionaLinha(grdBordero, True)
          
          Beep
          MsgBox "Borderô em utilização por outro usuário", vbInformation, Me.Caption
          'Desabilita seleção por linha
          Call SelecionaLinha(grdBordero)
          GoTo Sair
     End If
     
     'Inicia controle de tempo para informar Borderô Em Prova Zero
     sTempo = 0
     tmrAtualiza.Enabled = True
     
     'Carrega somente totais do borderô
     Call LoadGridBordero(True)
     
     Call LoadGridDeposito
     'Habilita seleção por linha
     Call SelecionaLinha(grdDataDeposito, True)
     
     Call LoadGridCheques
     'Habilita seleção por linha
     Call SelecionaLinha(grdCheques, True)
     
     Call LoadGridTotalBordero
     
Sair:
     'Seta ponteiro para default
     Screen.MousePointer = vbDefault

     Exit Sub
     
Erro_grdBordero_Click:

     'Rollback na transação
     g_cMainConnection.RollbackTrans
     
     Beep
     MsgBox Err.Description, vbCritical, Me.Caption

     'Desabilita seleção por linha
     Call SelecionaLinha(grdBordero)

     GoTo Sair

End Sub

Private Sub GrdBordero_DblClick()

Dim IdBordero As Long, iCancel As enumRetornoModal

On Error GoTo Erro_grdBordero_DblClick

     If grdBordero.Row <= 0 Then
          Exit Sub
     End If

     'Obtem o IdBordero do Bordero selecionado
     IdBordero = grdBordero.RowData(grdBordero.Row)
     
     'Setar parâmetros do form (bordero)
     Bordero.SetIdbordero (CStr(IdBordero))
     iCancel = Bordero.ShowModal(IdBordero, CStr(grdBordero.Text), True)


     'Se Encerrou com modificação, altera grid´s de Data e Cheque
     If iCancel = eRetornoOK Then
          'Muda ponteiro do windows
          Screen.MousePointer = vbHourglass
               
          'Leitura das novas informações pertencentes ao borderô
          Set rsBordero = g_cMainConnection.Execute(Proc_Selecionar.GetNumBordero(Geral.DataProcessamento, IdBordero))
          
          If rsBordero.EOF Then
               MsgBox "Erro na leitura das informações pertencentes ao borderô ", vbInformation, Me.Caption
               
               'Tenta retornar o status do borderô para a situação de (4)Prova Zero
               Call VoltaStatusBordero(grdBordero, Ambiente.iBor_RowSel)
               
               'Força Atualização do grid de borderô
               cmdAtualizaTela_Click
               
               'Seta ponteiro para default
               Screen.MousePointer = vbDefault
               Exit Sub
          End If
          
          'Atualiza totalizadores do borderô
          With grdBordero
               .Col = Coluna.Bor_QtdCheques: .Text = rsBordero!SomaQuantidade
               Acumulador.lQtdTotBordero = rsBordero!SomaQuantidade
               .Col = Coluna.Bor_VlrCheques: .Text = (CDbl(rsBordero!SomaValor) / 100)
               Acumulador.dVlrTotBordero = CDbl(rsBordero!SomaValor) / 100
          End With
               
          'Habilita seleção por linha
          Call SelecionaLinha(grdBordero, True)
          
          'Carrega somente totais do borderô
          Call LoadGridBordero(True)
          
          Call LoadGridDeposito
          'Habilita seleção por linha
          Call SelecionaLinha(grdDataDeposito, True)
          
          Call LoadGridCheques
          'Habilita seleção por linha
          Call SelecionaLinha(grdCheques, True)
          
          Call LoadGridTotalBordero
          
     End If

Sair:
     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_grdBordero_DblClick:
     
     Beep
     MsgBox Err.Description, vbCritical, Me.Caption
     GoTo Sair
     
End Sub

Private Sub grdBordero_GotFocus()
     
     'Posiciona o foco no grid CHEQUES, (Não pode haver foco
     'neste controle devido ao problema do KeyPress )
     grdCheques.SetFocus

End Sub

Private Sub grdBordero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then KeyAscii = 0
End Sub

Private Sub grdCheques_DblClick()

Dim IdCheque As Double, IdBordero As Long, iCancel As enumRetornoModal
Dim sDataDeposito As String, i As Integer, bMesmaDataDep As Boolean, bMesmoCheque As Boolean
Dim lngDataDeposito As Long

If grdCheques.Row <= 0 Then
     Exit Sub
End If

'Obtem o IdBordero do Bordero selecionado
IdBordero = grdBordero.RowData(grdBordero.Row)

'Obtem o IdCheque do Cheque selecionado
IdCheque = grdCheques.RowData(grdCheques.Row)

'Obtem a Data de depósito do grid
grdDataDeposito.Col = Coluna.Dep_Data
sDataDeposito = grdDataDeposito.Text
lngDataDeposito = FormataAMD(sDataDeposito)

'Habilita seleção por linha
Call SelecionaLinha(grdDataDeposito, True)

'Setar parâmetros do form (CHEQUE)
Cheque.SetIdCheque (IdCheque)
iCancel = Cheque.ShowModal(IdBordero, lngDataDeposito, IdCheque)

'Se Encerrou com modificação, altera grid´s de Data e Cheque
If iCancel = eRetornoOK Then
     'Muda ponteiro do windows
     Screen.MousePointer = vbHourglass
     
     grdCheques.Rows = 1
     Call SelecionaLinha(grdCheques)
     
     'Carrega grid de Depósito do borderô
     Call LoadGridDeposito

     bMesmaDataDep = False
     For i = 1 To grdDataDeposito.Rows - 1
          grdDataDeposito.Row = i
          If grdDataDeposito.Text = sDataDeposito Then
               bMesmaDataDep = True
               Exit For
          End If
     Next
     
     'Posiciona na mesma Data de Depósito do grid (Data Deposito)
     If bMesmaDataDep Then
          'Habilita seleção por linha
          grdDataDeposito.Row = i
          Call SelecionaLinha(grdDataDeposito, True)
          'Carrega grid com cheques referentes a mesma data de depósito
          Call LoadGridCheques
     
     Else
          If grdDataDeposito.Rows = 1 Then
               'grdDataDeposito.Row = 1
               'Desabilita seleção por linha
               Call SelecionaLinha(grdDataDeposito)
          Else
               'grdDataDeposito.Row = 2
               'Habilita seleção por linha
               Call SelecionaLinha(grdDataDeposito, True)
          End If
     End If
     
     'Apresenta frame com totais do Borderô
     Call LoadGridTotalBordero

     'Posiciona no mesmo registro do cheque do grid (CHEQUES)
     bMesmoCheque = False
     For i = 1 To grdCheques.Rows - 1
          If grdCheques.RowData(i) = IdCheque Then
               grdCheques.Row = i
               bMesmoCheque = True
               Exit For
          End If
     Next
     
     If bMesmoCheque Then
          'Habilita seleção por linha
          Call SelecionaLinha(grdCheques, True)
          Exit Sub
     Else
          If grdCheques.Rows = 1 Then
                
               'grdCheques.Row = 1
               'Desabilita seleção por linha
               Call SelecionaLinha(grdCheques)
          Else
       
               'grdCheques.Row = 2
               'Habilita seleção por linha
               Call SelecionaLinha(grdCheques, True)
          End If
     End If
End If

'Muda ponteiro do windows para default
Screen.MousePointer = vbDefault

End Sub

Private Sub grdDataDeposito_Click()

Dim lRetorno As Long

On Error GoTo Erro_grdDataDeposito_Click

     'Verifica se seleção do borderô já efetuada
     If Ambiente.iDep_RowSel = grdDataDeposito.Row Then Exit Sub
     
     'Seta ponteiro para aguardo..
     Screen.MousePointer = vbHourglass
     
     'Habilita seleção por linha
     Call SelecionaLinha(grdDataDeposito, True)

     Call LoadGridCheques
     'Habilita seleção por linha
     Call SelecionaLinha(grdCheques, True)
     
     Call LoadGridTotalBordero
     
Sair:
     'Seta ponteiro para default
     Screen.MousePointer = vbDefault

     Exit Sub
     
Erro_grdDataDeposito_Click:

     Beep
     MsgBox Err.Description, vbCritical, Me.Caption

     'Desabilita seleção por linha
     Call SelecionaLinha(grdBordero)

     GoTo Sair

End Sub

Private Sub grdDataDeposito_GotFocus()
     
     'Posiciona o foco no grid DEPOSITO, (Não pode haver foco
     'neste controle devido ao problema do KeyPress )
     grdCheques.SetFocus

End Sub

Private Sub SelecionaLinha(ByVal objeto As Object, Optional ByVal HabilitaSelecao As Boolean = False)

Dim ColunaFinal     As Integer

'Se habilita linha e não existe linha no grid, não habilitar
If HabilitaSelecao And objeto.Rows <= 1 Then
     objeto.HighLight = flexHighlightNever
     Exit Sub
End If

ColunaFinal = Switch(objeto.Name = "grdBordero", Coluna.Bor_NrBord, _
                    objeto.Name = "grdDataDeposito", Coluna.Dep_VlrDifer, _
                    objeto.Name = "grdCheques", Coluna.Chq_VlrCheque)

'Muda situação da Linha
With objeto
     
     'Se selecionado, guarda em variavel de ambiente o número da linha do selecionada
     If objeto.Name = "grdBordero" Then
          .Col = Coluna.Bor_NrBord  'Informar coluna do inicio do grid
          Ambiente.iBor_RowSel = IIf(HabilitaSelecao, .Row, 0)
     ElseIf objeto.Name = "grdDataDeposito" Then
          .Col = Coluna.Dep_Data  'Informar coluna do inicio do grid
          Ambiente.iDep_RowSel = IIf(HabilitaSelecao, .Row, 0)
     ElseIf objeto.Name = "grdCheques" Then
          .Col = Coluna.Chq_NrBanco  'Informar coluna do inicio do grid
          Ambiente.iChq_RowSel = IIf(HabilitaSelecao, .Row, 0)
     End If
     
     .HighLight = IIf(HabilitaSelecao, flexHighlightAlways, flexHighlightNever)
     .ColSel = ColunaFinal
     
End With

End Sub
Private Sub LoadGridBordero(Optional ByVal SomenteTotal As Boolean = False)

Dim i               As Integer
Dim sMinPendente    As String
Dim lRetorno        As Long

On Error GoTo Erro_LoadGridBordero
    
     If Not SomenteTotal Then
          'Seta ponteiro para aguardo..
          Screen.MousePointer = vbHourglass
          
          'Limpa variaveis do módulo
          Acumulador.lQtdTotCheques = 0
          Acumulador.lQtdTotBordero = 0
          Acumulador.dVlrTotCheques = 0
          Acumulador.dVlrTotBordero = 0
     
          ''''''''''''''''''''''''''''''''''''''''''''''
          '         Carrega grid com Borderôs          '
          ''''''''''''''''''''''''''''''''''''''''''''''
          'Limpa grid´s de Borderô, Depósito e Cheques
          grdBordero.Rows = 1
          Call LimpaTela

          'Obtem minuto máximo para identificar se capa de borderô perdida com status Em Prova Zero (G)
          sMinPendente = "0:" & Right("00" & CStr(g_Parametros.TMP_Pendente / 100), 2)
          
          Set rsBordero = g_cMainConnection.Execute(Proc_Selecionar.GetBorderoProvaZeroBloqueado(Geral.DataProcessamento, sMinPendente))
          If rsBordero.EOF Then
               cmdAtualizaTela.Enabled = False
               If Not AguardaDocumento Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
               Else
                    Set rsBordero = g_cMainConnection.Execute(Proc_Selecionar.GetBorderoProvaZeroBloqueado(Geral.DataProcessamento, sMinPendente))
                    If rsBordero.EOF Then
                         MsgBox "Não foi possível verificar se existe borderô para a data de movimento (" & DataProcessamento & ").", vbInformation, Me.Caption
                         Screen.MousePointer = vbDefault
                         Exit Sub
                    End If
               End If
               cmdAtualizaTela.Enabled = True
          End If
          
          'Desabilita seleção por linha
          Call SelecionaLinha(grdBordero)
          
          For i = 1 To rsBordero.RecordCount
               With grdBordero
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = Coluna.Bor_NrBord:     .Text = rsBordero!Num_Bordero
                    .Col = Coluna.Bor_QtdCheques: .Text = rsBordero!SomaQuantidade
                    .Col = Coluna.Bor_VlrCheques: .Text = CDbl(rsBordero!SomaValor) / 100
                    .Col = Coluna.Bor_Supervisor: .Text = False
                    .RowData(.Row) = rsBordero!IdBordero
               End With

               'Altera Status do borderô de (G)Em Prova Zero Bloqueado para (4)Para Prova Zero
               If rsBordero!Status = "G" Then
                    Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, rsBordero!IdBordero, "4"), lRetorno, adCmdText)
               End If
               
               rsBordero.MoveNext
          Next
          
          'Posiciona no primeiro registro do borderô
          grdBordero.Row = 1
     End If
     
     'Obtem a Qtd e o Valor Total  do borderô
     grdBordero.Col = Coluna.Bor_QtdCheques: Acumulador.lQtdTotBordero = grdBordero.Text
     grdBordero.Col = Coluna.Bor_VlrCheques: Acumulador.dVlrTotBordero = grdBordero.Text
     
     grdBordero.Col = Coluna.Bor_NrBord

     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_LoadGridBordero:

     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Beep
     MsgBox Err.Description & " (CarregaGrid)", vbCritical, Me.Caption
     Unload Me

End Sub
Private Sub LoadGridDeposito()
     
Dim i As Integer

Dim dVlrDeposito    As Double
Dim lQtdDeposito    As Long
Dim dVlrCheques     As Double
Dim lQtdCheques     As Long


On Error GoTo Erro_LoadGridDeposito

     'Limpa variaveis do módulo
     Acumulador.lQtdTotCheques = 0
     Acumulador.dVlrTotCheques = 0
     
     'Inicializa variáveis
     dVlrDeposito = 0
     lQtdDeposito = 0
     dVlrCheques = 0
     lQtdCheques = 0
     
     'Seta ponteiro para aguardo..
     Screen.MousePointer = vbHourglass
     
     'Limpa Tela
     Call LimpaTela
     Call SelecionaLinha(grdBordero, True)
     
     grdDataDeposito.Rows = 1
     grdCheques.Rows = 1
     
     Call SelecionaLinha(grdDataDeposito)
     Call SelecionaLinha(grdCheques)
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '         Carrega grid com Datas de Depósito           '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     Set RsDeposito = g_cMainConnection.Execute(Proc_Selecionar.GetDataDeposito(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row)))
     If RsDeposito.EOF Then
          MsgBox "Não existe depósito para o borderô na data (" & DataProcessamento & ").", vbInformation, Me.Caption
          Screen.MousePointer = vbDefault
          Exit Sub
     End If
     
     For i = 1 To RsDeposito.RecordCount
          With grdDataDeposito
               .Rows = .Rows + 1
               .Row = .Rows - 1
               .Col = Coluna.Dep_Data:            .Text = FormataDMA(RsDeposito!DataDeposito)
               .Col = Coluna.Dep_QtdDeposito:     .Text = RsDeposito!QuantidadeCheques
               .Col = Coluna.Dep_VlrDeposito:     .Text = RsDeposito!ValorDeposito
          End With
          RsDeposito.MoveNext
     Next
     
     'Carrega Soma dos cheques para calculo da diferença por data de depósito
     With grdDataDeposito
          For i = 1 To .Rows - 1
               .Row = i
               .Col = Coluna.Dep_Data
               Set RsDeposito = g_cMainConnection.Execute(Proc_Selecionar.GetSomatoriaCheques(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), FormataAMD(.Text)))
               
               'Obtem Qtd e Vlr. de depósito
               .Col = Coluna.Dep_QtdDeposito: lQtdDeposito = .Text
               .Col = Coluna.Dep_VlrDeposito: dVlrDeposito = .Text
               
               'Obtem o total (QTD e VLR) de cheques na data de depósito
               If Not RsDeposito.EOF Then
                    .Col = Coluna.Dep_QtdTabCheque:    .Text = IIf(IsNull(RsDeposito(0).Value), 0, RsDeposito(0).Value)
                    lQtdCheques = .Text
                    .Col = Coluna.Dep_VlrTabCheque:    .Text = IIf(IsNull(RsDeposito(1).Value), 0, RsDeposito(1).Value)
                    dVlrCheques = .Text
               Else
                    .Col = Coluna.Dep_QtdTabCheque:    .Text = 0
                    .Col = Coluna.Dep_VlrTabCheque:    .Text = 0
                    lQtdCheques = 0
                    dVlrCheques = 0
               End If
               
               'Acumula total de Qtd. e Valor para apresentar em Total do Borderô
               Acumulador.lQtdTotCheques = Acumulador.lQtdTotCheques + lQtdCheques
               Acumulador.dVlrTotCheques = Acumulador.dVlrTotCheques + dVlrCheques
               
               'Calcula diferença entre valor de depósito e valor total de cheques
               .Col = Coluna.Dep_QtdDifer:   .Text = Formato(lQtdCheques - lQtdDeposito, "I")
               .Col = Coluna.Dep_VlrDifer:   .Text = Formato(dVlrCheques - dVlrDeposito)
               
               'Colorir a coluna do grid conforme Qtd. e Valor
               Call ColorirColuna(grdDataDeposito, Coluna.Dep_QtdDifer)
               Call ColorirColuna(grdDataDeposito, Coluna.Dep_VlrDifer)
          Next
     
          If .Rows > 8 Then
               .ColWidth(Coluna.Dep_Data) = 920
          Else
               .ColWidth(Coluna.Dep_Data) = 1100
          End If

          'Posiciona no primeiro registro (DATA) do Depósito
          .Row = 1
          .Col = Coluna.Dep_Data
     
     End With
     
     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Exit Sub
     
Erro_LoadGridDeposito:

     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Beep
     MsgBox Err.Description & " (CarregaGrid)", vbCritical, Me.Caption
     Unload Me

End Sub
Private Sub LoadGridCheques()

Dim i As Integer

     'Seta ponteiro para aguardo..
     Screen.MousePointer = vbHourglass

     ''''''''''''''''''''''''''''''''''''''''''''''
     '         Carrega grid com Cheques           '
     ''''''''''''''''''''''''''''''''''''''''''''''
     If grdDataDeposito.Rows > 1 Then

          Set RsCheque = g_cMainConnection.Execute(Proc_Selecionar.GetCheques(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), FormataAMD(grdDataDeposito.Text)))
          
          'Limpa Grid
          grdCheques.Rows = 1
          
          For i = 1 To RsCheque.RecordCount
               With grdCheques
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = Coluna.Chq_NrSeq:      .Text = .Row
                    .Col = Coluna.Chq_NrBanco:    .Text = Left(RsCheque!CMC7, 3)
                    .Col = Coluna.Chq_NrAgencia:  .Text = Mid(RsCheque!CMC7, 4, 4)
                    .Col = Coluna.Chq_NrConta:    .Text = Mid(RsCheque!CMC7, 23, 7)
                    .Col = Coluna.Chq_NrCheque:   .Text = Mid(RsCheque!CMC7, 12, 6)
                    .Col = Coluna.Chq_CPF_CNPJ:   .Text = FormataCpfCnpj(RsCheque!CNPJCPF)
                    .Col = Coluna.Chq_VlrCheque:  .Text = Formato(RsCheque!Valor)
                    .RowData(.Row) = RsCheque!IdCheque
               End With
               RsCheque.MoveNext
          Next

          If grdCheques.Rows > 13 Then
               grdCheques.ColWidth(Coluna.Chq_CPF_CNPJ) = 1820
          Else
               grdCheques.ColWidth(Coluna.Chq_CPF_CNPJ) = 2000
          End If
          
          If grdCheques.Rows > 1 Then
               'Posiciona no primeiro registro do Depósito
               grdCheques.Row = 1
               grdCheques.Col = 1
          End If
     End If

     'Acerta botôes
     cmdIncluirCheque.Enabled = IIf(Me.grdDataDeposito.Rows > 1, True, False)
     cmdExcluirCheque.Enabled = IIf(grdCheques.Rows > 1, True, False)
     cmdEnviaSupervisor.Enabled = True
     cmdEncerrarCapa.Enabled = True
     
     'Seta ponteiro para default
     Screen.MousePointer = vbDefault

     Exit Sub
     
Erro_LoadGridDeposito:

     'Seta ponteiro para default
     Screen.MousePointer = vbDefault
     
     Beep
     MsgBox Err.Description & " (CarregaGrid)", vbCritical, Me.Caption
     Unload Me

End Sub
Private Sub LoadGridTotalBordero()

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '         Carrega grid com totais do bordero           '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     lblQtdChequesCalculado = Formato(Acumulador.lQtdTotCheques, "I")
     lblQtdChequesInformado = Formato(Acumulador.lQtdTotBordero, "I")
     lblVlrChequesCalculado = Formato(Acumulador.dVlrTotCheques)
     lblVlrChequesInformado = Formato(Acumulador.dVlrTotBordero)

End Sub
Private Sub VoltaStatusBordero(ByVal objeto As Object, ByVal LinhaAnterior As Integer)

Dim lRetorno As Long

Dim iLinhaAtual As Integer

On Error GoTo Erro_VoltaStatusBordero

     With objeto
          iLinhaAtual = .Row
          .Row = LinhaAnterior
          
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '         Atualiza status do Borderô para  Prova Zero            '
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          g_cMainConnection.BeginTrans
          
          Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBorderoDePara(Geral.DataProcessamento, .RowData(.Row), "4", "G"), lRetorno, adCmdText)
          
          g_cMainConnection.CommitTrans
     
          'Finaliza controle de tempo para informar Borderô Em Prova Zero
          sTempo = 0
          tmrAtualiza.Enabled = False
     
          .Row = iLinhaAtual
     End With

     Exit Sub
     
Erro_VoltaStatusBordero:

     'Rollback na transação
     g_cMainConnection.RollbackTrans
     
     Beep
     MsgBox Err.Description, vbCritical, Me.Caption

End Sub


Private Sub LimpaTela()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Limpar o grid de borderô antes da chamada desta sub    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     'Acerta seleção de linha do grid Borderô
     Call SelecionaLinha(grdBordero)
     
     'Acerta seleção de linha do grid Depósito
     grdDataDeposito.Rows = 1
     Call SelecionaLinha(grdDataDeposito)
     
     'Acerta seleção de linha do grid Cheques
     grdCheques.Rows = 1
     Call SelecionaLinha(grdCheques)
     
     'Limpa frame de totais do borderô
     lblQtdChequesInformado.Caption = ""
     lblVlrChequesInformado.Caption = ""
     lblQtdChequesCalculado.Caption = ""
     lblVlrChequesCalculado.Caption = ""

     'Acerta botôes
     cmdIncluirCheque.Enabled = False
     cmdExcluirCheque.Enabled = False
     cmdEnviaSupervisor.Enabled = False
     cmdEncerrarCapa.Enabled = False


End Sub
Private Sub tmrAtualiza_Timer()

Dim rsHoraAtual As ADODB.Recordset

     tmrAtualiza.Enabled = False
    
     If grdBordero.RowData(grdBordero.Row) <> 0 Then
          sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)

          If sTempo + Int(tmrAtualiza.Interval / 1000) >= g_Parametros.TMP_Pendente Then

               'Obs.: Utiliza a tabela StatusBordero apenas para obter a hora atual do servidor
               Set rsHoraAtual = g_cMainConnection.Execute("select distinct time() from PARAMETRO")

               'Atualizar a hora do bordero
               Call g_cMainConnection.Execute(Proc_atualizar.AtualizaHoraAtualBordero(Geral.DataProcessamento, grdBordero.RowData(grdBordero.Row), rsHoraAtual(0)))

               sTempo = 0
          End If
     End If
    
     tmrAtualiza.Enabled = True
    
     Set rsHoraAtual = Nothing
End Sub

Private Function AguardaDocumento() As Boolean

Dim sMinPendente As String

     Screen.MousePointer = vbDefault
    
     AguardaDocumento = False
     
     Set objAguardaDocumento = New AguardaDocumento
    
     'Obtem minuto máximo para identificar se capa de borderô perdida com status Em Prova Zero (G)
     sMinPendente = "0:" & Right("00" & CStr(g_Parametros.TMP_Pendente / 100), 2)

     objAguardaDocumento.SetConnection g_cMainConnection
     objAguardaDocumento.Tempo = (g_Parametros.TMP_Pendente / 10)
     objAguardaDocumento.SetStatusBar Me.StatusBar
     objAguardaDocumento.SetProgressBar Me.ProgressBar1
     objAguardaDocumento.SQL = Proc_Selecionar.GetBorderoProvaZeroBloqueado(Geral.DataProcessamento, sMinPendente, True)

     objAguardaDocumento.SetStatus "Aguardando borderô para prova zero ..."
     Do While Not objAguardaDocumento.ExisteDocumento()
        DoEvents
          objAguardaDocumento.SQL = Proc_Selecionar.GetBorderoProvaZeroBloqueado(Geral.DataProcessamento, sMinPendente, True)
     Loop
    
     If Not objAguardaDocumento.Finalizado Then
          If objAguardaDocumento.Recordset(0).Value > 0 Then
               AguardaDocumento = True
          End If
     End If
    
     Set objAguardaDocumento = Nothing
     Screen.MousePointer = vbDefault
    
End Function
