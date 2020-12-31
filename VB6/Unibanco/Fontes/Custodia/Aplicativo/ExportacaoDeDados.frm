VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ExportacaoDeDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação de dados"
   ClientHeight    =   7488
   ClientLeft      =   564
   ClientTop       =   744
   ClientWidth     =   10872
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7488
   ScaleWidth      =   10872
   Begin VB.Frame frmExportacao 
      Caption         =   " Layout para Exportação de Dados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   288
      TabIndex        =   20
      Top             =   120
      Width           =   10284
      Begin VB.Frame fraArquivoGeracao 
         Caption         =   " Arquivo para geração"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1644
         Left            =   4968
         TabIndex        =   21
         Top             =   288
         Width           =   5004
         Begin VB.TextBox txtArquivo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   192
            MaxLength       =   20
            TabIndex        =   7
            Top             =   528
            Width           =   2820
         End
         Begin VB.CommandButton cmdDiretorio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   4368
            Picture         =   "ExportacaoDeDados.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1248
            Width           =   468
         End
         Begin VB.TextBox txtDiretorio 
            Enabled         =   0   'False
            Height          =   288
            Left            =   192
            TabIndex        =   9
            Top             =   1224
            Width           =   4068
         End
         Begin VB.Label lblArquivo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nome do arquivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Left            =   192
            TabIndex        =   6
            Top             =   288
            Width           =   2820
         End
         Begin VB.Label lblDiretorio 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diretório"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Left            =   192
            TabIndex        =   8
            Top             =   984
            Width           =   4068
         End
      End
      Begin VB.ComboBox cmbFimLinha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2472
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1584
         Width           =   1692
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   324
         Left            =   8448
         TabIndex        =   17
         Top             =   6816
         Width           =   1524
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir LayOut"
         Height          =   324
         Left            =   6720
         TabIndex        =   16
         Top             =   6816
         Width           =   1524
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar Arquivo"
         Height          =   324
         Left            =   5016
         TabIndex        =   15
         Top             =   6816
         Width           =   1524
      End
      Begin VB.PictureBox picAdicionar 
         AutoRedraw      =   -1  'True
         Height          =   456
         Left            =   4344
         Picture         =   "ExportacaoDeDados.frx":00EA
         ScaleHeight     =   408
         ScaleWidth      =   408
         TabIndex        =   18
         Top             =   4032
         Width           =   456
      End
      Begin VB.PictureBox picRemover 
         Height          =   468
         Left            =   4344
         Picture         =   "ExportacaoDeDados.frx":052C
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   19
         Top             =   4824
         Width           =   468
      End
      Begin VB.ComboBox cmbDelimitador 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "ExportacaoDeDados.frx":096E
         Left            =   2472
         List            =   "ExportacaoDeDados.frx":0970
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1176
         Width           =   1692
      End
      Begin VB.TextBox txtRotulo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   300
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   3852
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   4284
         Left            =   4968
         TabIndex        =   14
         Top             =   2424
         Width           =   5004
         _ExtentX        =   8827
         _ExtentY        =   7557
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView 
         Height          =   4284
         Left            =   300
         TabIndex        =   12
         Top             =   2424
         Width           =   3876
         _ExtentX        =   6837
         _ExtentY        =   7557
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1032
         Top             =   4392
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ExportacaoDeDados.frx":0972
               Key             =   "ArquivoFechado"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ExportacaoDeDados.frx":0A6C
               Key             =   "Cheque"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ExportacaoDeDados.frx":0C06
               Key             =   "ArquivoAberto"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFimLinha 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fim de linha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   300
         TabIndex        =   4
         Top             =   1584
         Width           =   2088
      End
      Begin VB.Label lblCamposExportacao 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campos para exportação"
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
         Left            =   4968
         TabIndex        =   13
         Top             =   2160
         Width           =   5004
      End
      Begin VB.Label lblCamposArquivo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campos de arquivos"
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
         Left            =   300
         TabIndex        =   11
         Top             =   2160
         Width           =   3852
      End
      Begin VB.Label lblDelimitador 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delimitador de campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   300
         TabIndex        =   2
         Top             =   1176
         Width           =   2088
      End
      Begin VB.Label lblRotulo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome do Rótulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   0
         Top             =   360
         Width           =   3852
      End
   End
End
Attribute VB_Name = "ExportacaoDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Proc_Selecionar           As New Custodia.Selecionar
Dim Proc_Inserir              As New Custodia.Inserir
Dim Proc_Excluir              As New Custodia.Excluir
Dim aStructFile()             'Elementos (Nome Campo, Tamanho, Alinhamento, Zeros, Ocorrência)
Dim aAlinhamento()
Dim aZeros()
Dim bAlterouEstrutura         As Boolean          'Identificador de alteração no layout
Dim bAlterouComplemento       As Boolean          'Identificador de alteração de complementos do layout

Public m_Col_Tamanho          As Integer
Public m_EscolhaExportacao    As String           'Identifica qual arquivos à exportar (B)orderô ou (I)nstrução

'Constantes para referencia de coluna do ListView
Private Const c_Col_Ordenacao = 0, c_Col_Nome = 1, c_Col_Tamanho = 2
Private Const c_Col_Alinhamento = 3, c_Col_Zeros = 4
'Constantes para referencia do combo de alinhamento
Private Const c_AlignNone = 0, c_AlignLeft = 1, c_AlignRight = 2
'Constantes para referencia do combo de preenchimento com zeros
Private Const c_ZerosNone = 0, c_Zerosleft = 1, c_ZerosRight = 2
'Constantes para referencia de aStructFile()
Private Const c_Stru_Nome = 0, c_Stru_Tamanho = 1, c_Stru_Alinham = 2, c_Stru_Zeros = 3
'

Private Sub cmbDelimitador_Click()

     Dim iCount     As Integer
     Dim iSelected  As Integer
     
     If Not Me.cmbDelimitador.Visible Then Exit Sub
     If ListView.ListItems.Count = 0 Then Exit Sub
     
     bAlterouComplemento = True
     iSelected = Me.ListView.SelectedItem.Index
     
     If cmbDelimitador.ListIndex > 0 And ListView.ListItems.Count > 0 Then
          bAlterouEstrutura = True
          With ListView
               For iCount = 1 To .ListItems.Count
                    Set .SelectedItem = .ListItems(iCount)
                    .SelectedItem.SubItems(c_Col_Alinhamento) = Alinhamento(c_AlignNone)
               Next
               Set .SelectedItem = .ListItems(iSelected)
          End With
     Else
          With ListView
               For iCount = 1 To .ListItems.Count
                    Set .SelectedItem = .ListItems(iCount)
                    If .SelectedItem.SubItems(c_Col_Alinhamento) = Alinhamento(c_AlignNone) Then
                         bAlterouEstrutura = True
                         .SelectedItem.SubItems(c_Col_Alinhamento) = Alinhamento(c_AlignLeft)
                    End If
               Next
               Set .SelectedItem = .ListItems(iSelected)
          End With
     
     End If

End Sub

Private Sub cmbDelimitador_KeyPress(KeyAscii As Integer)

     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
     End If
     
End Sub

Private Sub cmbFimLinha_Click()

     bAlterouComplemento = True
     
End Sub

Private Sub cmbFimLinha_KeyPress(KeyAscii As Integer)

     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
     End If

End Sub

Private Sub cmdDiretorio_Click()
    
     bAlterouComplemento = True
     
     Me.txtDiretorio.Text = Parametros.BrowseForFolder(Me, camMyComputer)
    
End Sub

Private Sub cmdGerar_Click()

On Error GoTo Err_cmdGerar_Click

     Dim bBeginTrans     As Boolean
     Dim rsResult        As New ADODB.Recordset
     Dim lRetorno        As Long
     Dim strDataServ     As String
     Dim iCount          As Integer
     Dim bFinalizou      As Boolean
     
     bFinalizou = False
     
     'Consiste todos campos necessários para exportação de dados
     If Trim(txtArquivo) = "" Then
          Beep
          MsgBox "Favor informar o nome do arquivo para exportação de dados", vbInformation, Me.Caption
          txtArquivo.SetFocus
          GoTo Sair
     End If
     If Trim(txtDiretorio) = "" Then
          Beep
          MsgBox "Favor informar o diretório para exportação de dados", vbInformation, Me.Caption
          cmdDiretorio.SetFocus
          GoTo Sair
     End If
     If ListView.ListItems.Count = 0 Then
          Beep
          MsgBox "Não é possível gerar arquivo sem a informação de campos para exportação", vbInformation, Me.Caption
          ListView.SetFocus
          GoTo Sair
     End If
     
     
     Screen.MousePointer = vbHourglass

     If bAlterouEstrutura Or bAlterouComplemento Then
          If bAlterouComplemento Then
               Set rsResult = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lRetorno, adCmdText)
               If lRetorno <> 0 Then
                    Beep
                    MsgBox "Problema na geração do relatório de inconsistências!", vbCritical, Me.Caption
                    GoTo Sair
               End If
               
               strDataServ = Format(rsResult!Data, "yyyymmdd")
          End If
          
          'Abre transação
          bBeginTrans = True
          g_cMainConnection.BeginTrans
               
          'Remove antiga estrutura do layout
          If bAlterouEstrutura Then Call g_cMainConnection.Execute(Proc_Excluir.RemoveCamposExportacao(m_EscolhaExportacao), lRetorno, adCmdText)
          If bAlterouComplemento Then Call g_cMainConnection.Execute(Proc_Excluir.RemoveComplementoExportacao(m_EscolhaExportacao), lRetorno, adCmdText)
          
          'Adiciona nova estrutura do layout
          If bAlterouComplemento Then
               Call g_cMainConnection.Execute(Proc_Inserir.InsereComplementoExportacao( _
                                                  m_EscolhaExportacao, _
                                                  CLng(strDataServ), _
                                                  IIf(Trim(txtRotulo) = "", "", txtRotulo), _
                                                  cmbDelimitador.ListIndex, _
                                                  cmbFimLinha.ListIndex, _
                                                  Geral.UsuarioLogin, _
                                                  Trim(txtArquivo), _
                                                  Trim(txtDiretorio)), lRetorno, adCmdText)
          End If
          
          If bAlterouEstrutura Then
               With ListView
                    For iCount = 1 To .ListItems.Count
                         Set .SelectedItem = .ListItems(iCount)
                         Call g_cMainConnection.Execute(Proc_Inserir.InsereCamposExportacao( _
                                                            m_EscolhaExportacao, _
                                                            .SelectedItem, _
                                                            .SelectedItem.SubItems(c_Col_Nome), _
                                                            .SelectedItem.SubItems(c_Col_Tamanho), _
                                                            Alinhamento(.SelectedItem.SubItems(c_Col_Alinhamento)), _
                                                            Zeros(.SelectedItem.SubItems(c_Col_Zeros)), _
                                                            Left(.SelectedItem.Key, 3)), lRetorno, adCmdText)
                                                                                
                    Next
               End With
          End If
          
          'Encerra transação
          g_cMainConnection.CommitTrans
     
     End If
     
     bFinalizou = GerarArquivo
     
Sair:
    
     If Not (rsResult Is Nothing) Then Set rsResult = Nothing
     Screen.MousePointer = vbDefault
     If bFinalizou Then cmdSair_Click
     Exit Sub

Err_cmdGerar_Click:
     Beep
     Screen.MousePointer = vbDefault
     
     'Cancela transação
     If bBeginTrans Then g_cMainConnection.RollbackTrans

     MsgBox "Não foi possível gerar exportação de dados" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Sair

End Sub

Private Sub cmdImprimir_Click()

On Error GoTo Err_cmdImprimir_Click

     Dim rsCampos        As New ADODB.Recordset
     Dim rsComplementos  As New ADODB.Recordset
     
     Screen.MousePointer = vbHourglass
     
     If bAlterouEstrutura Or bAlterouComplemento Then
          Beep
          MsgBox "O Layout foi alterado, favor gerar arquivo para poder imprimir com as modificações.", vbInformation, Me.Caption
          GoTo Exit_cmdImprimir_Click
     End If
     
     Set rsCampos = g_cMainConnection.Execute(Proc_Selecionar.GetCamposExportacao(m_EscolhaExportacao))
     
     If rsCampos.EOF Then
          Beep
          Screen.MousePointer = vbDefault

          MsgBox "Não existe layout à ser impresso" & vbCrLf & vbCrLf & _
               "Após montagem dos campos para exportação, gerar " & vbCrLf & _
               "arquivo para que seja finalizada a definição do layout ", vbInformation, Me.Caption
          GoTo Exit_cmdImprimir_Click
     End If
     
     Set rsComplementos = g_cMainConnection.Execute(Proc_Selecionar.GetComplementoExportacao(m_EscolhaExportacao))
     
     If rsComplementos.EOF Then
          Beep
          Screen.MousePointer = vbDefault
          MsgBox "Não foi possível carregar informações do layout, tente gerar arquivo antes de imprimir", vbCritical, Me.Caption
          GoTo Exit_cmdImprimir_Click
     End If

     With Principal.CrystalReport
          .ReportFileName = App.path & "\Reports\RelExportacao.rpt"
          .SelectionFormula = "{ComplementoExportacao.Referencia}='" & m_EscolhaExportacao & "' and {CamposExportacao.Referencia}= '" & m_EscolhaExportacao & "'"
          .Formulas(0) = "OpcaoImpressao = '" & Switch(m_EscolhaExportacao = "B", "( Borderô )", m_EscolhaExportacao = "D", "( Alteração de Data )", m_EscolhaExportacao = "C", "( Cheques Baixados )", m_EscolhaExportacao = "O", "( Cheques Data Boa )") & "'"
          .WindowState = crptMaximized
          .WindowTitle = "Emissão do Layout de exportação de dados"
          .Action = 0
     
     End With
     

Exit_cmdImprimir_Click:
     Screen.MousePointer = vbDefault
     
     If Not (rsCampos Is Nothing) Then Set rsCampos = Nothing
     If Not (rsComplementos Is Nothing) Then Set rsComplementos = Nothing
     Exit Sub
     
Err_cmdImprimir_Click:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível gerar exportação de dados" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_cmdImprimir_Click

End Sub

Private Sub cmdSair_Click()
     
     Unload Me

End Sub

Private Sub Form_Activate()

     'Centraliza form
     Me.Left = (Screen.Width - Me.Width) / 2
     Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     If KeyCode = vbKeyEscape Then Unload Me
     
End Sub

Private Sub Form_Load()
     
     
     If m_EscolhaExportacao = "B" Then
          Me.Caption = Me.Caption & " - Borderô"
     ElseIf m_EscolhaExportacao = "D" Then
          Me.Caption = Me.Caption & " - Alteração de Data"
     ElseIf m_EscolhaExportacao = "C" Then
          Me.Caption = Me.Caption & " - Cheques Baixados"
     ElseIf m_EscolhaExportacao = "O" Then
          Me.Caption = Me.Caption & " - Cheques Data Boa"
     End If
     
     aAlinhamento = Array("Não", "Esquerda", "Direita")
     aZeros = Array("Não", "Esquerda", "Direita")
     
     m_Col_Tamanho = c_Col_Tamanho
     
     If m_EscolhaExportacao = "B" Then
          Call CarregaCamposTreeView_Bor
     ElseIf m_EscolhaExportacao = "D" Then
          Call CarregaCamposTreeView_AltDT
     ElseIf m_EscolhaExportacao = "C" Then
          Call CarregaCamposTreeView_CHQ
     ElseIf m_EscolhaExportacao = "O" Then
          Call CarregaCamposTreeView_BOA
     End If
     
     Call CarregaCombos
     Call IniciaListView
     Call CarregaTela
     
     'Identificador de alteração no layout
     bAlterouEstrutura = False
     bAlterouComplemento = False
     
End Sub
Private Sub CarregaCamposTreeView_Bor()

On Error GoTo Err_CarregaCamposTreeView_Bor
    
     Dim rsBordero       As New ADODB.Recordset
     Dim rsDataDeposito  As New ADODB.Recordset
     Dim RsCheque        As New ADODB.Recordset
     Dim nd              As Node
     Dim iFields         As Integer
     Dim iRelativo       As Integer
     Dim iIndexStruct    As Integer

     Screen.MousePointer = vbHourglass
     
     'Busca somente as informações estruturais da tabela Borderô
     Set rsBordero = g_cMainConnection.Execute(Proc_Selecionar.GetBorderoConfirmacao(19000101))
     
     'Busca somente as informações estruturais da tabela Data Depósito
     Set rsDataDeposito = g_cMainConnection.Execute(Proc_Selecionar.GetDataDepositoBordero(19000101, 0, 19000101))
     
     'Busca somente as informações estruturais da tabela Cheques
     Set RsCheque = g_cMainConnection.Execute(Proc_Selecionar.GetEstruturaCheque(19000101))
     
     
     TreeView.Nodes.Clear
     iIndexStruct = 0
     'Gera ocorrência(0) para campo Rótulo
     ReDim Preserve aStructFile(3, iIndexStruct)
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campo para Rótulo         '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Rotulo", "Rótulo", "ArquivoFechado", "ArquivoAberto")
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Borderô  '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Bordero", "Arquivo Bordero", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (rsBordero.Fields.Count - 1)
          'Elimina campos sem interesse para exportação
          If InStr("idbordero*status*horaatual", LCase(rsBordero(iFields).Name)) = 0 Then
               Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "BOR_" & rsBordero(iFields).Name, rsBordero(iFields).Name, "Cheque")
               
               iIndexStruct = iIndexStruct + 1
               ReDim Preserve aStructFile(3, iIndexStruct)
               aStructFile(c_Stru_Nome, iIndexStruct) = "BOR_" & rsBordero(iFields).Name
               
               If rsBordero(iFields).Precision = 255 Then
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsBordero(iFields).DefinedSize
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
               Else
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsBordero(iFields).Precision
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
               End If
               
               aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone

          End If
     Next
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Data Depósito      '
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "DataDeposito", "Arquivo Data Depósito", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (rsDataDeposito.Fields.Count - 1)
          'Elimina campos sem interesse para exportação
          If InStr("dataprocessamento*idbordero*status", LCase(rsDataDeposito(iFields).Name)) = 0 Then
               Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "DTD_" & rsDataDeposito(iFields).Name, rsDataDeposito(iFields).Name, "Cheque")
               iIndexStruct = iIndexStruct + 1
               ReDim Preserve aStructFile(3, iIndexStruct)
               aStructFile(c_Stru_Nome, iIndexStruct) = "DTD_" & rsDataDeposito(iFields).Name
               
               If rsDataDeposito(iFields).Precision = 255 Then
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsDataDeposito(iFields).DefinedSize
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
               Else
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsDataDeposito(iFields).Precision
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
               End If
               
               aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone
          End If
     Next
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Cheque   '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Cheque", "Arquivo Cheque", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (RsCheque.Fields.Count - 1)
          If InStr("dataprocessamento*idbordero*idcheque*possuierro*status", LCase(RsCheque(iFields).Name)) = 0 Then
               Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "CHQ_" & RsCheque(iFields).Name, RsCheque(iFields).Name, "Cheque")
               iIndexStruct = iIndexStruct + 1
               ReDim Preserve aStructFile(3, iIndexStruct)
               aStructFile(c_Stru_Nome, iIndexStruct) = "CHQ_" & RsCheque(iFields).Name
               
               If RsCheque(iFields).Precision = 255 Then
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).DefinedSize
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
               Else
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).Precision
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
               End If
               
               aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone
          End If
     Next

Exit_CarregaCamposTreeView_Bor:
     Screen.MousePointer = vbDefault
     If Not (rsBordero Is Nothing) Then Set rsBordero = Nothing
     If Not (rsDataDeposito Is Nothing) Then Set rsDataDeposito = Nothing
     If Not (RsCheque Is Nothing) Then Set RsCheque = Nothing
     Exit Sub
     
Err_CarregaCamposTreeView_Bor:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível carregar campos das tabelas" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_CarregaCamposTreeView_Bor

End Sub

Private Sub IniciaListView()
     
     Dim ColHead         As ColumnHeaders
     Dim lsitem          As ListItem
     
     '''''''''''''''''''''''''''''''''''''''''''''
     Set ListView.DropHighlight = Nothing
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'Finaliza o icone de drag, (Mais utilizado quando .Drag vbBeginDrag)
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     ListView.Drag vbEndDrag
     
     ListView.GridLines = True
     ListView.ListItems.Clear
     ListView.ColumnHeaders.Add
     Set ColHead = ListView.ColumnHeaders
     If ListView.ColumnHeaders.Count < 1 Then ListView.ColumnHeaders.Add
     With ColHead(1): .Width = ListView.Width * 0.1: .Text = "Seq":  End With
     If ListView.ColumnHeaders.Count < 2 Then ListView.ColumnHeaders.Add
     With ColHead(2): .Width = ListView.Width * 0.35: .Text = "Campo": .Alignment = lvwColumnLeft: End With
     If ListView.ColumnHeaders.Count < 3 Then ListView.ColumnHeaders.Add
     With ColHead(3): .Width = ListView.Width * 0.19: .Text = "Tamanho": .Alignment = lvwColumnRight: End With
     If ListView.ColumnHeaders.Count < 4 Then ListView.ColumnHeaders.Add
     With ColHead(4): .Width = ListView.Width * 0.19: .Text = "Alinhar": .Alignment = lvwColumnCenter: End With
     If ListView.ColumnHeaders.Count < 5 Then ListView.ColumnHeaders.Add
     With ColHead(5): .Width = ListView.Width * 0.17: .Text = "Zeros": .Alignment = lvwColumnCenter: End With
     
     
End Sub

Private Sub ListView_DblClick()

     Dim iCount As Integer
     
     If ListView.ListItems.Count = 0 Then Exit Sub

     Load ExportacaoPropriedades
     With ExportacaoPropriedades
          .SetForm Me
          'Seleciona option de alinhamento na tela de complemento da exportacao
          Select Case Alinhamento(ListView.SelectedItem.SubItems(c_Col_Alinhamento))
               Case c_AlignNone
                    .optAlinha_Sem.Value = True
               Case c_AlignLeft
                    .optAlinha_Esquerdo.Value = True
               Case c_AlignRight
                    .optAlinha_Direito.Value = True
          End Select
          'Seleciona option de preenchimento com zeros na tela de complemento da exportacao
          Select Case Zeros(ListView.SelectedItem.SubItems(c_Col_Zeros))
               Case c_ZerosNone
                    .optZeros_Sem.Value = True
               Case c_Zerosleft
                    .optZeros_Esquerda.Value = True
               Case c_ZerosRight
                    .optZeros_Direita.Value = True
          End Select
          
          .Caption = "Propriedades de (" & Me.ListView.SelectedItem.SubItems(c_Col_Nome) & ")"

          For iCount = 0 To UBound(aStructFile, 2)
               If aStructFile(c_Stru_Nome, iCount) = ListView.SelectedItem.Key Then
                   .txtTamanho.Tag = aStructFile(c_Stru_Tamanho, iCount)
                    Exit For
               End If

          Next

          .Show vbModal, Me
     
          If Not .GetCancelou Then
               'Identificador de alteração no layout
               bAlterouEstrutura = True
               ListView.SelectedItem.SubItems(c_Col_Tamanho) = Val(.txtTamanho.Text)
               ListView.SelectedItem.SubItems(c_Col_Alinhamento) = aAlinhamento(Switch(.optAlinha_Sem, c_AlignNone, .optAlinha_Esquerdo, c_AlignLeft, .optAlinha_Direito, c_AlignRight))
               ListView.SelectedItem.SubItems(c_Col_Zeros) = aZeros(Switch(.optZeros_Sem, c_ZerosNone, .optZeros_Esquerda, c_Zerosleft, .optZeros_Direita, c_ZerosRight))
          End If
     
     End With
     Unload ExportacaoPropriedades
     
End Sub

Private Sub picAdicionar_Click()
     
     Dim iCount          As Integer
     Dim lsitem          As ListItem
     Dim iIndexStruct    As Integer
     
     If Not (TreeView.SelectedItem Is Nothing) Then
          If InStr("bordero*datadeposito*cheque*rotulo*alteracao*databoa", LCase(TreeView.SelectedItem.Key)) Then
               Beep
               MsgBox "Favor escolher um campo pertecente à pasta do arquivo", vbInformation, Me.Caption
               Exit Sub
          End If
     
          For iCount = 1 To ListView.ListItems.Count
               If ListView.ListItems(iCount).Key = TreeView.SelectedItem.Key Then
                    Beep
                    MsgBox "Campo de tabela já selecionado, favor escolher outro", vbInformation, Me.Caption
                    Exit Sub
               End If
          Next

          iCount = ListView.ListItems.Count + 1
          Set lsitem = ListView.ListItems.Add(, TreeView.SelectedItem.Key, CStr(iCount))
          iIndexStruct = PesquisaIndexStruct(TreeView.SelectedItem.Key)
          lsitem.SubItems(c_Col_Nome) = TreeView.SelectedItem.Text
          
          If Left(TreeView.SelectedItem.Key, 4) = "ROT_" Then
               lsitem.SubItems(c_Col_Tamanho) = Len(Trim(txtRotulo))
          Else
               lsitem.SubItems(c_Col_Tamanho) = aStructFile(c_Stru_Tamanho, iIndexStruct)
          End If
          
          If Me.cmbDelimitador.ListIndex = 0 Then
               lsitem.SubItems(c_Col_Alinhamento) = aAlinhamento(aStructFile(c_Stru_Alinham, iIndexStruct))
               lsitem.SubItems(c_Col_Zeros) = aZeros(aStructFile(c_Stru_Zeros, iIndexStruct))
          Else
               lsitem.SubItems(c_Col_Alinhamento) = Alinhamento(c_AlignNone)
               lsitem.SubItems(c_Col_Zeros) = Zeros(c_ZerosNone)
          End If

          Set ListView.SelectedItem = ListView.ListItems(ListView.ListItems.Count)
          
          'Identificador de alteração no layout
          bAlterouEstrutura = True
          
          Exit Sub
     Else
          Beep
          MsgBox "Favor selecionar um dos campos na janela de 'Campos de arquivos'", vbInformation, Me.Caption
     End If
     Me.ListView.SelectedItem = False

End Sub
Private Sub CarregaCombos()

     'Carrega combo delimitador
     cmbDelimitador.AddItem "Não"
     cmbDelimitador.AddItem """"
     cmbDelimitador.AddItem "'"
     cmbDelimitador.AddItem ":"
     cmbDelimitador.AddItem ";"
     cmbDelimitador.ListIndex = 0
     
     'Carrega combo Fim de Linha
     cmbFimLinha.AddItem "Nenhum"
     cmbFimLinha.AddItem "Line Feed"
     cmbFimLinha.AddItem "Return"
     cmbFimLinha.AddItem "Return/Line Feed"
     cmbFimLinha.ListIndex = 0
     
End Sub

Private Function PesquisaIndexStruct(ByVal sKEY As String) As Integer

     Dim iCount As Integer
     
     PesquisaIndexStruct = 0
     
     For iCount = 0 To UBound(aStructFile, 2)
     
          If aStructFile(c_Stru_Nome, iCount) = sKEY Then
               PesquisaIndexStruct = iCount
               Exit Function
          End If
     Next

End Function

Private Sub picRemover_Click()
     
     Dim iCount          As Integer
     Dim iAtualItem      As Integer
     
     If ListView.SelectedItem Is Nothing Then
          If ListView.ListItems.Count = 0 Then Exit Sub
          Beep
          MsgBox "Favor selecionar um dos ítens na janela de 'Campos para Exportação'", vbInformation, Me.Caption
          Exit Sub
     End If

     iAtualItem = ListView.SelectedItem
     
     'Remove linha do ListView
     ListView.ListItems.Remove CVar(ListView.SelectedItem.Key)
     'Renomeia a sequência das linhas do ListView
     For iCount = 1 To ListView.ListItems.Count
          ListView.ListItems(iCount).Text = iCount
     Next
     
     If iAtualItem > ListView.ListItems.Count Then iAtualItem = iAtualItem - 1
     If iAtualItem = 1 And ListView.ListItems.Count = 1 Then iAtualItem = 1
     If iAtualItem = 1 And ListView.ListItems.Count < 1 Then iAtualItem = 0

     If iAtualItem > 0 Then Set ListView.SelectedItem = ListView.ListItems(iAtualItem)
     
End Sub

Private Sub TreeView_Collapse(ByVal Node As MSComctlLib.Node)

     Node.Image = Me.ImageList1.ListImages("ArquivoFechado").Key

End Sub

Private Sub TreeView_DblClick()
     
     picAdicionar_Click
     
End Sub

Public Function Alinhamento(ByVal vntParam As Variant) As Variant

     If IsNumeric(vntParam) Then
          Alinhamento = aAlinhamento(vntParam)
     Else
          Alinhamento = Switch(aAlinhamento(c_AlignNone) = vntParam, c_AlignNone, _
                               aAlinhamento(c_AlignLeft) = vntParam, c_AlignLeft, _
                               aAlinhamento(c_AlignRight) = vntParam, c_AlignRight)
     End If

End Function
Public Function Zeros(ByVal vntParam As Variant) As Variant
     
     If IsNumeric(vntParam) Then
          Zeros = aZeros(vntParam)
     Else
          Zeros = Switch(aZeros(c_ZerosNone) = vntParam, c_ZerosNone, _
                         aZeros(c_Zerosleft) = vntParam, c_Zerosleft, _
                         aZeros(c_ZerosRight) = vntParam, c_ZerosRight)
     End If

End Function

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)

     Node.Image = Me.ImageList1.ListImages("ArquivoAberto").Key

End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)

     If InStr("rotulo*bordero*datadeposito*cheque*alteracao*databoa", LCase(Node.Key)) > 0 Then
          If Node.Expanded Then
               Node.Expanded = False
          Else
               Node.Expanded = True
          End If
     End If

End Sub

Private Sub txtArquivo_KeyPress(KeyAscii As Integer)
     
     If KeyAscii = vbKeyEscape Then Exit Sub
     
     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
          Exit Sub
     End If
     
     If InStr("^~;:.,<>'´`?/|\+-*! ", Chr(KeyAscii)) Then
          KeyAscii = 0
          Exit Sub
     End If
     
     bAlterouComplemento = True
     If KeyAscii = vbKeyBack Then Exit Sub
     
     If Len(txtArquivo) >= txtArquivo.MaxLength Then
          Beep
          KeyAscii = 0
          MsgBox "Número máximo permitido é de " & CStr(txtArquivo.MaxLength) & " caracteres", vbInformation, Me.Caption
          txtArquivo.SelStart = 0
          txtArquivo.SelLength = txtArquivo.MaxLength
          Exit Sub
     End If

End Sub

Private Sub txtRotulo_KeyPress(KeyAscii As Integer)

     txtRotulo.Tag = ""
     
     If KeyAscii = vbKeyEscape Then Exit Sub
     
     txtRotulo.Tag = "CALL"
     
     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
          Exit Sub
     End If
     
     bAlterouComplemento = True
     If KeyAscii = vbKeyBack Then Exit Sub
     
     If Len(Me.txtRotulo) >= txtRotulo.MaxLength Then
          Beep
          KeyAscii = 0
          MsgBox "Número máximo permitido é de " & CStr(txtRotulo.MaxLength) & " caracteres", vbInformation, Me.Caption
          txtRotulo.SelStart = 0
          txtRotulo.SelLength = txtRotulo.MaxLength
          Exit Sub
     End If

End Sub
Private Sub InsereAlteraRotuloListView()

     Dim iCount As Integer
     Dim iLstCount As Integer
     Dim nd As Node

     For iCount = 1 To TreeView.Nodes.Count
          'Posiciona no indice referente a pasta do Rótulo
          TreeView.SelectedItem = TreeView.Nodes(iCount)
          
          'Verifica se o Nó é pasta ou campo
          If TreeView.Nodes(iCount) = "Rótulo" Then
          
               If Trim(txtRotulo) <> "" Then
                    'Altera ou insere campo
                    If TreeView.Nodes.Item(iCount).Children = 0 Then
                         Set nd = TreeView.Nodes.Add(iCount, tvwChild, "ROT_" & Trim(txtRotulo), Trim(txtRotulo), "Cheque", "Cheque")
                    Else
                         'Posiciona no indice referente ao campo do Rótulo
                         TreeView.Nodes(iCount).Child = Trim(txtRotulo)
                         'Altera do ListView a descrição do rótulo
                         For iLstCount = 1 To ListView.ListItems.Count
                              If Left(ListView.ListItems(iLstCount).Key, 4) = "ROT_" Then
                                   ListView.ListItems(iLstCount).Key = "ROT_" & Trim(txtRotulo)
                                   ListView.ListItems(iLstCount).ListSubItems(c_Col_Nome) = Trim(txtRotulo)
                                   ListView.ListItems(iLstCount).ListSubItems(c_Col_Tamanho) = Len(Trim(txtRotulo.Text))
                                   Exit For
                              End If
                         Next
                         
                    End If
                    Exit For
               Else
                    'Exclui campo do treeview
                    If TreeView.Nodes.Item(iCount).Children > 0 Then
                         TreeView.Nodes.Remove (TreeView.Nodes(iCount).Child.Index)
                    End If
                    
                    For iLstCount = 1 To ListView.ListItems.Count
                         If Left(ListView.ListItems(iLstCount).Key, 4) = "ROT_" Then
                              Set ListView.SelectedItem = ListView.ListItems(iLstCount)
                              picRemover_Click
                              Exit For
                         End If
                    Next
                    
               End If
               Exit For
          End If
               
     Next
     
End Sub
Private Sub CarregaTela()

On Error GoTo Err_CarregaTela

     Dim rsCampos        As New ADODB.Recordset
     Dim rsComplementos  As New ADODB.Recordset
     Dim lsitem          As ListItem
     Dim iCount          As Integer
     
     Screen.MousePointer = vbHourglass
     
     Set rsCampos = g_cMainConnection.Execute(Proc_Selecionar.GetCamposExportacao(m_EscolhaExportacao))
     
     If rsCampos.EOF Then GoTo Exit_CarregaTela
     
     Set rsComplementos = g_cMainConnection.Execute(Proc_Selecionar.GetComplementoExportacao(m_EscolhaExportacao))
     
     If rsComplementos.EOF Then
          Beep
          MsgBox "Não foi possível carregar informações do layout" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
          GoTo Exit_CarregaTela
     End If

     'Carrega complementos do layout de exportação
     txtRotulo.Text = IIf(IsNull(rsComplementos!Rotulo), "", rsComplementos!Rotulo)
     txtArquivo.Text = rsComplementos!NomeArquivo
     txtDiretorio.Text = rsComplementos!Diretorio
     cmbDelimitador.ListIndex = rsComplementos!Delimitador
     cmbFimLinha.ListIndex = rsComplementos!FimDeLinha
     
     'Carrega campos do layout de exportação
     If Trim(txtRotulo) <> "" Then Call InsereAlteraRotuloListView

     While Not rsCampos.EOF
          iCount = ListView.ListItems.Count + 1
          Set lsitem = ListView.ListItems.Add(, rsCampos!SiglaTabela & "_" & Trim(rsCampos!nome), CStr(iCount))
          lsitem.SubItems(c_Col_Nome) = Trim(rsCampos!nome)
          
          lsitem.SubItems(c_Col_Tamanho) = rsCampos!Tamanho
               
          lsitem.SubItems(c_Col_Alinhamento) = aAlinhamento(rsCampos!Alinhamento)
          lsitem.SubItems(c_Col_Zeros) = aZeros(rsCampos!Zeros)
          
          rsCampos.MoveNext
     Wend
     
     Set ListView.SelectedItem = ListView.ListItems(1)

Exit_CarregaTela:
     Screen.MousePointer = vbDefault
     If Not (rsCampos Is Nothing) Then Set rsCampos = Nothing
     If Not (rsComplementos Is Nothing) Then Set rsComplementos = Nothing
     Exit Sub

Err_CarregaTela:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível carregar informações do layout" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_CarregaTela

End Sub

Private Sub txtRotulo_Validate(Cancel As Boolean)
     
     If txtRotulo.Tag = "CALL" Then
          Call InsereAlteraRotuloListView
     End If
     
End Sub

Private Function GerarArquivo() As Boolean

GerarArquivo = False

On Error GoTo Err_GerarArquivo

     Dim iCount     As Integer
     Dim sArquivo   As String
     Dim rsResult   As New ADODB.Recordset
     Dim lRetorno   As Long
     Dim sSql       As String
     
     Dim bTabDTD    As Boolean     'Indentificador da existencia de campo da tabela Data Depósito
     Dim bTabCHQ    As Boolean     'Indentificador da existencia de campo da tabela Cheque
     Dim bTabBAI    As Boolean     'Identificaador da existencia de campo da tabela Cheques Baixados
     Dim bTabALT    As Boolean     'Identificaador da existencia de campo da tabela Alteração de Data
     
     Dim aComplCampo()             'Configuração para formatação do campo (Tamanho, Alinhamento, Zeros)
     Dim iIndexConf As Integer
     Dim sLinha     As String      'Linha para geração do registro do arquivo
          
     'Monta instrução sql para seleção de campos conforme layout definido pelo usuário
     bTabDTD = False: bTabCHQ = False
     bTabBAI = False: bTabALT = False
     sSql = "Select "
     
     iIndexConf = 0
     With ListView
          .HideSelection = True
          For iCount = 1 To .ListItems.Count

               .SelectedItem = .ListItems(iCount)
               'Seleciona somente os children´s
               If InStr("BOR_DTD_CHQ_ROT_BAI_ALT_BOA_", Left(.SelectedItem.Key, 4)) Then
                    If Left(.SelectedItem.Key, 4) = "ROT_" Then
                         sSql = sSql & "'" & Trim(.SelectedItem.SubItems(c_Col_Nome)) & "' as Rotulo"
                    ElseIf Left(.SelectedItem.Key, 4) = "BOR_" Then
                         sSql = sSql & "BORD" & "." & Trim(.SelectedItem.SubItems(c_Col_Nome))
                    Else
                         sSql = sSql & Left(.SelectedItem.Key, 3) & "." & Trim(.SelectedItem.SubItems(c_Col_Nome))
                    End If
                    
                    iIndexConf = iIndexConf + 1
                    ReDim Preserve aComplCampo(2, iIndexConf)
                    aComplCampo(0, iIndexConf) = .SelectedItem.SubItems(c_Col_Alinhamento)
                    aComplCampo(1, iIndexConf) = .SelectedItem.SubItems(c_Col_Zeros)
                    aComplCampo(2, iIndexConf) = .SelectedItem.SubItems(c_Col_Tamanho)
                    
                    If m_EscolhaExportacao = "B" Then
                         'Identificador de campos para exportacao de Borderôs
                         If Left(.SelectedItem.Key, 4) = "DTD_" Then bTabDTD = True
                         If Left(.SelectedItem.Key, 4) = "CHQ_" Then bTabCHQ = True
                    Else
                         'Identificador de campos para exportacao de Instrução
                         If Left(.SelectedItem.Key, 4) = "BAI_" Then bTabBAI = True
                         If Left(.SelectedItem.Key, 4) = "ALT_" Then bTabALT = True
                    End If
                    
                    If iCount <> .ListItems.Count Then sSql = sSql & ", "
               End If
          Next
     End With
     
     If m_EscolhaExportacao = "B" Then
          ''''''''''''''''''''''''''''''''''''''''''''''
          '    Exportação das Tabelas de Borderô       '
          ''''''''''''''''''''''''''''''''''''''''''''''
          sSql = sSql & " from Bordero BORD "
                    
          If bTabDTD And bTabCHQ Then
               sSql = sSql & " Left Join (DataDeposito DTD"
               sSql = sSql & " Left Join Cheque CHQ"
               sSql = sSql & " On  DTD.dataprocessamento = CHQ.dataprocessamento"
               sSql = sSql & " And DTD.idbordero = CHQ.idbordero"
               sSql = sSql & " And DTD.datadeposito = CHQ.datadeposito)"
               sSql = sSql & " On  BORD.dataprocessamento = DTD.dataprocessamento"
               sSql = sSql & " And BORD.idbordero = DTD.idbordero"
               sSql = sSql & " Where CHQ.DataProcessamento = " & Geral.DataProcessamento
               sSql = sSql & " And   BORD.Status = 'E'"           'Somente bordero Confirmado
               sSql = sSql & " And   DTD.Status = '1'"           'Somente Data OK
               sSql = sSql & " And   CHQ.Status = 'E'"           'Somente cheque Confirmado
          
          ElseIf bTabDTD Then
               sSql = sSql & " Left Join DataDeposito DTD"
               sSql = sSql & " On  BORD.dataprocessamento = DTD.dataprocessamento"
               sSql = sSql & " And BORD.idbordero = DTD.idbordero"
               sSql = sSql & " Where DTD.DataProcessamento = " & Geral.DataProcessamento
               sSql = sSql & " And   BORD.Status = 'E'"           'Somente bordero Confirmado
               sSql = sSql & " And   DTD.Status = '1'"           'Somente Data OK
               
          ElseIf bTabCHQ Then
               sSql = sSql & " Left Join Cheque CHQ"
               sSql = sSql & " On  BORD.dataprocessamento = CHQ.dataprocessamento"
               sSql = sSql & " And BORD.idbordero = CHQ.idbordero"
               sSql = sSql & " Where CHQ.DataProcessamento = " & Geral.DataProcessamento
               sSql = sSql & " And   BORD.Status = 'E'"           'Somente bordero Confirmado
               sSql = sSql & " And   CHQ.Status = 'E'"           'Somente cheque Confirmado
          Else
               sSql = sSql & " Where BORD.DataProcessamento = " & Geral.DataProcessamento
          End If
          
     ElseIf m_EscolhaExportacao = "D" Then
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          '    Exportação da Tabela de Alteração de Data    '
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          sSql = sSql & " From AlteracaoData ALT"
          sSql = sSql & " Where ALT.DataProcessamento = " & Geral.DataProcessamento
     ElseIf m_EscolhaExportacao = "C" Then
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          '    Exportação da Tabela de Cheques Baixados     '
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          sSql = sSql & " From ChequesBaixados BAI"
     ElseIf m_EscolhaExportacao = "O" Then
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          '    Exportação da Tabela de Movto Data Boa       '
          '''''''''''''''''''''''''''''''''''''''''''''''''''
          sSql = sSql & " From  ChequeDataBoa BOA"
          sSql = sSql & " Where BOA.DataProcessamento = " & Geral.DataProcessamento
          sSql = sSql & " And   BOA.Fusao = " & CInt(True)
     
     End If
     
     Set rsResult = g_cMainConnection.Execute(sSql, lRetorno, adCmdText)
     
     sArquivo = txtDiretorio & "\" & Trim(txtArquivo.Text) & ".txt"
     Kill (sArquivo)
     
     If rsResult.EOF Then
          Beep
          MsgBox "Não existe informações para exportação de dados. Verifique !", vbCritical, Me.Caption
          GoTo Exit_GerarArquivo
     End If
     
     Open sArquivo For Binary Access Write As #1
     
     'Inicializa Progress Bar
     With Principal
          .ProgressBar1.Min = 0
          .ProgressBar1.Max = rsResult.RecordCount
          .ProgressBar1.Visible = True
          .StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = ""
          .StatusBarPrincipal.Panels(StatusBar.Col_Descrição).Text = "Exportando arquivo de " & _
                    Switch(m_EscolhaExportacao = "B", "Borderô", m_EscolhaExportacao = "D", "Alter. Data", m_EscolhaExportacao = "C", "Cheques Baixados", m_EscolhaExportacao = "O", "Cheques Data Boa") & " ..."
     End With
     
     Do While Not rsResult.EOF
          
          sLinha = ""
          For iCount = 1 To rsResult.Fields.Count

               'Verifica se campos com delimitador
               If cmbDelimitador.ListIndex <> 0 Then sLinha = sLinha & cmbDelimitador.Text
               
               sLinha = sLinha + ConverteCampo(rsResult(iCount - 1), _
                                        rsResult(iCount - 1).Name, _
                                        ListView.ListItems(iCount).SubItems(c_Col_Tamanho), _
                                        Zeros(ListView.ListItems(iCount).SubItems(c_Col_Zeros)), _
                                        Alinhamento(ListView.ListItems(iCount).SubItems(c_Col_Alinhamento)))
               If cmbDelimitador.ListIndex <> 0 Then sLinha = sLinha & cmbDelimitador.Text
                    
               'Adiciona (,) para opção com delimitador
               If cmbDelimitador.ListIndex <> 0 Then
                    If iCount <> rsResult.Fields.Count Then
                         sLinha = sLinha & ","
                    End If
               End If
               
               If iCount = rsResult.Fields.Count Then
                    'Insere fim de linha
                    Select Case cmbFimLinha.ListIndex
                         Case 1
                              sLinha = sLinha & vbLf   'Line Feed
                         Case 2
                              sLinha = sLinha & vbCr   'Return
                         Case 3
                              sLinha = sLinha & vbCrLf 'Return + Line Feed
                    End Select
               End If
          Next
          
          'Grava registro
          Put #1, , sLinha


          
          'Atualiza Progress Bar
          Principal.ProgressBar1.Value = rsResult.AbsolutePosition
          Principal.StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = rsResult.AbsolutePosition & "/" & rsResult.RecordCount
          
          rsResult.MoveNext
     Loop

     Beep
     Screen.MousePointer = vbDefault
     
     MsgBox "Finalizado a exportação de dados", vbInformation + vbOKOnly, Me.Caption
     
     'Saida com sucesso na finalização
     GerarArquivo = True
     
Exit_GerarArquivo:
     ListView.HideSelection = False
     Screen.MousePointer = vbDefault
     If Not (rsResult Is Nothing) Then Set rsResult = Nothing
     Close
     With Principal
          .StatusBarPrincipal.Panels(StatusBar.Col_Descrição).Text = ""
          .ProgressBar1.Visible = False
          .StatusBarPrincipal.Panels(StatusBar.Col_ContadorProgressBar).Text = ""
     End With
     Exit Function
     
Err_GerarArquivo:
     'Verifica se tentou apagar o arquivo e este não existe
     If Err.Number = 53 Then Resume Next
     
     Beep
     Screen.MousePointer = vbDefault

     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório para geração do arquivo. Verifique !", vbCritical, Me.Caption
          GoTo Exit_GerarArquivo
     End If
     MsgBox "Não foi possível carregar informações do layout" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_GerarArquivo

End Function
Private Function ConverteCampo(ByVal vntCampo As Variant, _
                                ByVal sNomeCampo As String, _
                                ByVal iTamanho As Integer, _
                                ByVal iPreencheComZeros As Integer, _
                                ByVal iAlinhamento As Integer) As String

     Dim sCampo As String

     If IsNull(vntCampo) Then
          sCampo = ""
     Else
          sCampo = CStr(vntCampo)
          
          'Verifica se campo contém valor decimal
          If IsNumeric(vntCampo) Then
               If InStr(LCase(sNomeCampo), "valor") > 0 Or InStr(LCase(sNomeCampo), "vlr") Then
                    If iTamanho <= 19 Then
                        'Verifica se valor contém casas decimais, se não acrescenta 2 dígitos transformando em inteiro
                         If sCampo <> CStr(Fix(vntCampo)) Then
                              sCampo = CStr(vntCampo * 100)
                         Else
                              sCampo = sCampo & "00"
                         End If
                    End If
               End If
          End If
     End If
     
     If iPreencheComZeros > c_ZerosNone Then
          If iAlinhamento = c_AlignRight Then
               If iPreencheComZeros = c_Zerosleft Then
                    sCampo = Right(String(iTamanho, "0") + sCampo, iTamanho)
               Else
                    sCampo = Left(sCampo + String(iTamanho, "0"), iTamanho)
               End If
               
          ElseIf iAlinhamento = c_AlignLeft Then
               If iPreencheComZeros = c_Zerosleft Then
                    sCampo = Right(String(iTamanho, "0") + sCampo, iTamanho)
               Else
                    sCampo = Left(sCampo + String(iTamanho, "0"), iTamanho)
               End If
          Else
               If Len(sCampo) > iTamanho Then
                    sCampo = Left(sCampo, iTamanho)
               End If
          End If
     Else
          If iAlinhamento = c_AlignRight Then
               sCampo = Right(String(iTamanho, " ") + sCampo, iTamanho)
          ElseIf iAlinhamento = c_AlignLeft Then
               sCampo = Left(sCampo + String(iTamanho, " "), iTamanho)
          Else
               If Len(sCampo) > iTamanho Then
                    sCampo = Left(sCampo, iTamanho)
               End If
          End If
     End If

     ConverteCampo = sCampo

End Function
Private Sub CarregaCamposTreeView_CHQ()
     
On Error GoTo Err_CarregaCamposTreeView_CHQ

     Dim RsCheque        As New ADODB.Recordset        'Cheques baixados
     Dim nd              As Node
     Dim iFields         As Integer
     Dim iRelativo       As Integer
     Dim iIndexStruct    As Integer

     Screen.MousePointer = vbHourglass
     
     'Busca somente as informações estruturais da tabela Cheques Baixados
     Set RsCheque = g_cMainConnection.Execute(Proc_Selecionar.GetStructChequesBaixados)
     
     TreeView.Nodes.Clear
     iIndexStruct = 0
     'Gera ocorrência(0) para campo Rótulo
     ReDim Preserve aStructFile(3, iIndexStruct)
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campo para Rótulo         '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Rotulo", "Rótulo", "ArquivoFechado", "ArquivoAberto")
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Cheques Baixados   '
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Cheque", "Arquivo Cheques Baixados", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (RsCheque.Fields.Count - 1)
          Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "BAI_" & RsCheque(iFields).Name, RsCheque(iFields).Name, "Cheque")
          iIndexStruct = iIndexStruct + 1
          ReDim Preserve aStructFile(3, iIndexStruct)
          aStructFile(c_Stru_Nome, iIndexStruct) = "BAI_" & RsCheque(iFields).Name
          
          If RsCheque(iFields).Precision = 255 Then
               aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).DefinedSize
               aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
          Else
               aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).Precision
               aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
          End If
          
          aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone
     Next

Exit_CarregaCamposTreeView_CHQ:
     Screen.MousePointer = vbDefault
     If Not (RsCheque Is Nothing) Then Set RsCheque = Nothing
     Exit Sub
     
Err_CarregaCamposTreeView_CHQ:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível carregar campos das tabelas" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_CarregaCamposTreeView_CHQ

End Sub
Private Sub CarregaCamposTreeView_AltDT()
          
On Error GoTo Err_CarregaCamposTreeView_AltDT
    
     Dim rsAlteracao     As New ADODB.Recordset        'Alteração de data
     Dim nd              As Node
     Dim iFields         As Integer
     Dim iRelativo       As Integer
     Dim iIndexStruct    As Integer

     Screen.MousePointer = vbHourglass
     
     'Busca somente as informações estruturais da tabela Alteração de Data
     Set rsAlteracao = g_cMainConnection.Execute(Proc_Selecionar.GetStructAlteracaoData())
     
     TreeView.Nodes.Clear
     iIndexStruct = 0
     'Gera ocorrência(0) para campo Rótulo
     ReDim Preserve aStructFile(3, iIndexStruct)
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campo para Rótulo         '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Rotulo", "Rótulo", "ArquivoFechado", "ArquivoAberto")
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Alteração de Data  '
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Alteracao", "Arquivo Alteração de Data", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (rsAlteracao.Fields.Count - 1)
          'Elimina campos sem interesse para exportação
          If InStr("idalteracao", LCase(rsAlteracao(iFields).Name)) = 0 Then
               Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "ALT_" & rsAlteracao(iFields).Name, rsAlteracao(iFields).Name, "Cheque")
               
               iIndexStruct = iIndexStruct + 1
               ReDim Preserve aStructFile(3, iIndexStruct)
               aStructFile(c_Stru_Nome, iIndexStruct) = "ALT_" & rsAlteracao(iFields).Name
               
               If rsAlteracao(iFields).Precision = 255 Then
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsAlteracao(iFields).DefinedSize
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
               Else
                    aStructFile(c_Stru_Tamanho, iIndexStruct) = rsAlteracao(iFields).Precision
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
               End If
               
               aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone

          End If
     Next
     
Exit_CarregaCamposTreeView_AltDT:
     Screen.MousePointer = vbDefault
     If Not (rsAlteracao Is Nothing) Then Set rsAlteracao = Nothing
     Exit Sub
     
Err_CarregaCamposTreeView_AltDT:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível carregar campos das tabelas" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_CarregaCamposTreeView_AltDT

End Sub
Private Sub CarregaCamposTreeView_BOA()
     
On Error GoTo Err_CarregaCamposTreeView_BOA

     Dim RsCheque        As New ADODB.Recordset        'Cheques baixados
     Dim nd              As Node
     Dim iFields         As Integer
     Dim iRelativo       As Integer
     Dim iIndexStruct    As Integer

     Screen.MousePointer = vbHourglass
     
     'Busca somente as informações estruturais da tabela Cheques Data Boa
     Set RsCheque = g_cMainConnection.Execute(Proc_Selecionar.GetChequeDataBoa(19000101))
     
     TreeView.Nodes.Clear
     iIndexStruct = 0
     'Gera ocorrência(0) para campo Rótulo
     ReDim Preserve aStructFile(3, iIndexStruct)
                    aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
     
     '''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campo para Rótulo         '
     '''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "Rotulo", "Rótulo", "ArquivoFechado", "ArquivoAberto")
     
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     '    Adiciona campos da Tabela Cheques Baixados   '
     '''''''''''''''''''''''''''''''''''''''''''''''''''
     Set nd = TreeView.Nodes.Add(, , "DataBoa", "Arquivo Cheques Data Boa", "ArquivoFechado", "ArquivoAberto")
     iRelativo = TreeView.Nodes.Count
     For iFields = 0 To (RsCheque.Fields.Count - 1)
          Set nd = TreeView.Nodes.Add(iRelativo, tvwChild, "BOA_" & RsCheque(iFields).Name, RsCheque(iFields).Name, "Cheque")
          iIndexStruct = iIndexStruct + 1
          ReDim Preserve aStructFile(3, iIndexStruct)
          aStructFile(c_Stru_Nome, iIndexStruct) = "BOA_" & RsCheque(iFields).Name
          
          If RsCheque(iFields).Precision = 255 Then
               aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).DefinedSize
               aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignLeft
          Else
               aStructFile(c_Stru_Tamanho, iIndexStruct) = RsCheque(iFields).Precision
               aStructFile(c_Stru_Alinham, iIndexStruct) = c_AlignRight
          End If
          
          aStructFile(c_Stru_Zeros, iIndexStruct) = c_ZerosNone
     Next

Exit_CarregaCamposTreeView_BOA:
     Screen.MousePointer = vbDefault
     If Not (RsCheque Is Nothing) Then Set RsCheque = Nothing
     Exit Sub
     
Err_CarregaCamposTreeView_BOA:
     Beep
     Screen.MousePointer = vbDefault
     MsgBox "Não foi possível carregar campos das tabelas" + vbCrLf + vbCrLf + Err.Description, vbCritical, Me.Caption
     GoTo Exit_CarregaCamposTreeView_BOA

End Sub
