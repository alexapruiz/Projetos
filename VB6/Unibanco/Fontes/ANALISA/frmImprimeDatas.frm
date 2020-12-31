VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImprimeDatas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir datas processadas"
   ClientHeight    =   3528
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3528
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3252
      Top             =   2280
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
            Picture         =   "frmImprimeDatas.frx":0000
            Key             =   "FLD_CLOSE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImprimeDatas.frx":0112
            Key             =   "FLD_OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImprimeDatas.frx":0224
            Key             =   "COMPUTER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3312
      Left            =   96
      TabIndex        =   2
      Top             =   120
      Width           =   3036
      _ExtentX        =   5355
      _ExtentY        =   5842
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   396
      Left            =   3228
      TabIndex        =   1
      Top             =   660
      Width           =   972
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   396
      Left            =   3228
      TabIndex        =   0
      Top             =   180
      Width           =   972
   End
End
Attribute VB_Name = "frmImprimeDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_NumPag                        As Integer

Private Const ALTURA_RODA_PE = 500 'para impressao


Private Sub cmdFechar_Click()
    Unload Me
End Sub


Public Sub cmdImprimir_Click()

    Dim c                       As New PCabecalho
    Dim cnn                     As ADODB.Connection
    Dim sServidor               As String
    Dim sDataBase               As String
    Dim strUsuario              As String
    Dim strSenha                As String
    Dim sDataProcessamento      As Long
    Dim iID_Estacao             As Integer
    Dim sNomeEstacao            As String
    Dim sSql                    As String
    Dim rst                     As ADODB.Recordset
    Dim rstCMC7                 As ADODB.Recordset
    Dim rstCB                   As ADODB.Recordset
    Dim rstErros                As ADODB.Recordset
    Dim cmd                     As ADODB.Command
    Dim y                       As Long
    Dim sStr                    As String
    Dim Count                   As Integer
    Dim Porcentagem             As Currency
    
    Const LINHA = 300
    
'    printer.Height = Printer.ScaleHeight
'    printer.Width = Printer.ScaleWidth
'
'    printer.Show
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Pega a data de processamento e o ID_Estacao'
    '''''''''''''''''''''''''''''''''''''''''''''
    If Not TreeView1.SelectedItem Is Nothing Then
        If IsDate(TreeView1.SelectedItem.Text) Then
            MsgBox "É necessário selecionar uma estação para a impressão.", vbExclamation
            Exit Sub
        Else
            sDataProcessamento = Format(TreeView1.SelectedItem.Parent.Text, "YYYYMMDD")
            iID_Estacao = TreeView1.SelectedItem.Tag
            sNomeEstacao = TreeView1.SelectedItem.Text
        End If
    Else
        MsgBox "Selecione uma estação para a impressão.", vbExclamation
        Exit Sub
    End If
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set rstErros = New ADODB.Recordset
    Set rstCMC7 = New ADODB.Recordset
    Set rstCB = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    With cnn
        sServidor = PegarOpcaoINI("Conexao", "Servidor", "MDI_NT1")
        sDataBase = PegarOpcaoINI("Conexao", "DataBaseDestino", "Analisa")
        strUsuario = PegarOpcaoINI("Conexao", "Usuario", App.Path & "\MDI_Conexao.ini")
        strSenha = PegarOpcaoINI("Conexao", "Senha", App.Path & "\MDI_Conexao.ini")
        
        .Provider = "SQLOLEDB"
        .ConnectionTimeout = 5
        .CursorLocation = adUseClient
        .ConnectionString = "Server=" & sServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";Database=" & sDataBase
        .Open
    End With
    
    Set cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdStoredProc
    
    c.DataProcessamento = sDataProcessamento
    c.Estacao = sNomeEstacao
    c.Titulo = "ANALISE DE UTILIZAÇÃO DE SCANNER MC93"
    
    c.Imprimir
    
    
    m_NumPag = 0
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'não tenho tempo para terminar as classes de impressao'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'DEPOIS EU COLOCO NAS CLASSES
    
           sSql = "SELECT CAP.TempoCaptura,"
    sSql = sSql & "       CAP.TempoConfirmacao,"
    sSql = sSql & "       CAP.TempoTrocaLotes,"
    sSql = sSql & "       CAP.TempoResolucaoErros,"
    sSql = sSql & "       CAP.QtdeErros,"
    sSql = sSql & "       CAP.PercTempoErro,"
    sSql = sSql & "       CAP.QtdeLotesCapturados,"
    sSql = sSql & "       CAP.QtdeLotesProcessados,"
    sSql = sSql & "       CAP.QtdeLotesCancelados,"
    sSql = sSql & "       CAP.ProdutividadeCaptura,"
    sSql = sSql & "       CAP.ProdutividadeCapturaErro,"
    sSql = sSql & "       CAP.CMC7_Reconhecido,"
    sSql = sSql & "       CAP.CMC7_Erros,"
    sSql = sSql & "       CAP.CMC7_Porcent_Erros,"
    sSql = sSql & "       CAP.CB_Reconhecido,"
    sSql = sSql & "       CAP.CB_Erros,"
    sSql = sSql & "       CAP.CB_Porcent_Erros"
    sSql = sSql & "  FROM Captura CAP, Estacao EST"
    sSql = sSql & " WHERE CAP.DataProcessamento = " & sDataProcessamento
    sSql = sSql & "   AND CAP.ID_Estacao = EST.ID_Estacao"
    sSql = sSql & "   AND EST.ID_Estacao = " & iID_Estacao

    Set rst = cnn.Execute(sSql)
    
    '''''''''''''''''''
    'Comeca a imprimir'
    '''''''''''''''''''
    Printer.Font.Name = "Courier New"
    
    Printer.CurrentX = 0
    
    Printer.Print ""
    Printer.Print "Tempo de Captura          - " & rst!TempoCaptura
    Printer.Print "Tempo de Confirmacao      - " & rst!TempoConfirmacao
    Printer.Print "Tempo de Troca de Lotes   - " & rst!TempoTrocaLotes
    Printer.Print "Tempo de Solucao de Erros - " & rst!TempoResolucaoErros
    Printer.Print "% Tempo Erro x Captura    - " & rst!PercTempoErro
    Printer.Print ""
    Printer.Print "Qtde Lotes Capturados     - " & rst!QtdeLotesCapturados
    Printer.Print "Qtde Doctos Processados   - " & rst!QtdeLotesProcessados
    Printer.Print "Qtde Doctos Cancelados    - " & rst!QtdeLotesCancelados
    Printer.Print "Produtividade por Hora    - " & rst!ProdutividadeCaptura & " ( Captura )"
    Printer.Print "Produtividade por Hora    - " & rst!ProdutividadeCapturaErro & " ( Captura + Solução de Erros )"
    Printer.Print ""
    
    Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 15), , BF
    
    If rst!QtdeErros <> 0 Then
        Printer.Font.Name = "Times New Roman"
        Printer.Font.Size = "14"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Utilização do Scanner")) / 2
        Printer.Print "Utilização do Scanner"
        
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = 8
        
        Printer.Print ""
        Printer.Print "Qtde de Erros Ocorridos - " & rst!QtdeErros
        Printer.Print "Documentos por Erro     - " & rst!CMC7_Erros
        Printer.Print ""
        
        Printer.Print "Código do Erro     Descrição do Erro                           Quantidade    Tempo de Parada"
        
        cmd.CommandText = "GetErro"
        
        cmd.Parameters(1) = sDataProcessamento
        cmd.Parameters(2) = iID_Estacao
        
        '''''''''''''''''
        'Insere os Erros'
        '''''''''''''''''
        Set rstErros = cmd.Execute()
        
        Do While Not rstErros.EOF()
        
            y = Printer.CurrentY
        
            Printer.Print rstErros!Cod_Erro
            Printer.CurrentX = 1800: Printer.CurrentY = y
            Printer.Print rstErros!Descricao
            Printer.CurrentX = 6000: Printer.CurrentY = y
            Printer.Print rstErros!Qtde
            Printer.CurrentX = 7400: Printer.CurrentY = y
            Printer.Print rstErros!TempoParada
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se currentY + altura de uma linha na impressora + altura do roda pe'
            'maior que altura do papel, é nova pagina                           '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If y + (Printer.TextHeight("Altura de uma linha") * 2) + ALTURA_RODA_PE > Printer.ScaleHeight Then
            
                ImprimeRodaPe
            
                Printer.NewPage
                c.Imprimir
                Printer.CurrentX = 0
                Printer.Print "Código do Erro     Descrição do Erro                           Quantidade    Tempo de Parada"

            End If
        
            rstErros.MoveNext
        Loop
        rstErros.Close
        
    End If
    
    '''''''''''''''''''''''''
    'Insere os Erros de CMC7'
    '''''''''''''''''''''''''
    If rst!CMC7_Reconhecido <> 0 Then
        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 15), , BF
        
        Printer.Font.Name = "Times New Roman"
        Printer.Print ""
        Printer.Font.Size = "14"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Leitura de CMC-7")) / 2
        Printer.Print "Leitura de CMC-7"
        
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = 8
        
        Printer.Print ""
        Printer.Print "Total Reconhecido - " & rst!CMC7_Reconhecido
        Printer.Print "Total de Erros    - " & rst!CMC7_Erros
        Printer.Print "% de Erros        - " & rst!CMC7_Porcent_Erros
        Printer.Print ""
        
        Printer.Print "Documentos com Leitura Incorreta                               Quantidade    Porcentagem"
        
        cmd.CommandText = "GetCMC7_Erro"
        cmd.Parameters(1) = sDataProcessamento
        cmd.Parameters(2) = iID_Estacao
        
        Set rstCMC7 = cmd.Execute()
        
        Do While Not rstCMC7.EOF()
        
            y = Printer.CurrentY

            Printer.Print Left(rstCMC7!Descricao, 50)
            sStr = Trim(Left(rstCMC7!Descricao, 50))
            Count = 0
            Do While sStr = Trim(Left(rstCMC7!Descricao, 50))
                Count = Count + 1
                rstCMC7.MoveNext
                Porcentagem = (Count / rstCMC7.RecordCount) * 100
                If rstCMC7.EOF Then Exit Do
            Loop
            
            Printer.CurrentX = 6000: Printer.CurrentY = y
            Printer.Print Count
            Printer.CurrentX = 7400: Printer.CurrentY = y
            Printer.Print Format(Porcentagem, "00.00") & "%"
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se currentY + altura de uma linha na impressora + altura do roda pe'
            'maior que altura do papel, é nova pagina                           '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If y + (Printer.TextHeight("Altura de uma linha") * 2) + ALTURA_RODA_PE > Printer.ScaleHeight Then
            
                ImprimeRodaPe
            
                Printer.NewPage
                c.Imprimir
                Printer.CurrentX = 0
                Printer.Print "Código do Erro     Descrição do Erro                           Quantidade    Tempo de Parada"

            End If
            
        Loop
        
    End If
    
    
    '''''''''''''''''''''''
    'Insere os Erros de CB'
    '''''''''''''''''''''''
    If rst!CB_Reconhecido <> 0 Then
        'If rst!CMC7_Reconhecido <> 0 Then printer.newpage
        
        Printer.Line (0, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY + 15), , BF
        
        Printer.FontName = "Times New Roman"
        Printer.Print ""
        Printer.Font.Size = "14"
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Leitura de Codigo de Barras")) / 2
        Printer.Print "Leitura de Codigo de Barras"
        
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = 8
        
        Printer.Print ""
        Printer.Print "Total Reconhecido - " & rst!CB_Reconhecido
        Printer.Print "Total de Erros    - " & rst!CB_Erros
        Printer.Print "% de Erros        - " & rst!CB_Porcent_Erros
        Printer.Print ""
        
        Printer.Print "Documentos com Leitura Incorreta                               Quantidade    Porcentagem"
        
        cmd.CommandText = "GetCB_Erro"
        cmd.Parameters(1) = sDataProcessamento
        cmd.Parameters(2) = iID_Estacao
        
        Set rstCB = cmd.Execute()
        
        Do While Not rstCB.EOF()
        
            y = Printer.CurrentY

            Printer.Print Left(rstCB!Descricao, 50)
            sStr = Trim(Left(rstCB!Descricao, 50))
            Count = 0
            Do While sStr = Trim(Left(rstCB!Descricao, 50))
                Count = Count + 1
                rstCB.MoveNext
                Porcentagem = (Count / rstCB.RecordCount) * 100
                If rstCB.EOF Then Exit Do
            Loop
            
            Printer.CurrentX = 6000: Printer.CurrentY = y
            Printer.Print Count
            Printer.CurrentX = 7400: Printer.CurrentY = y
            Printer.Print Format(Porcentagem, "00.00") & "%"
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se currentY + altura de uma linha na impressora + altura do roda pe'
            'maior que altura do papel, é nova pagina                           '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If y + (Printer.TextHeight("Altura de uma linha") * 2) + ALTURA_RODA_PE > Printer.ScaleHeight Then
            
                ImprimeRodaPe
            
                Printer.NewPage
                c.Imprimir
                Printer.CurrentX = 0
                Printer.Print "Documentos com Leitura Incorreta                               Quantidade    Porcentagem"

            End If
            
        Loop
        
    End If
    
    Printer.EndDoc
    MsgBox "Enviando para a impressora.", vbInformation
    
    
    rst.Close
    cnn.Close

End Sub

Private Sub ImprimeRodaPe()
    
    m_NumPag = m_NumPag + 1

    '''''''''''''''
    'Faz uma linha'
    '''''''''''''''
    frmPrint.CurrentY = frmPrint.ScaleHeight - ALTURA_RODA_PE
    frmPrint.Line (0, frmPrint.CurrentY)-(frmPrint.ScaleWidth, frmPrint.CurrentY)
    
    frmPrint.CurrentY = frmPrint.CurrentY + 50
    frmPrint.CurrentX = frmPrint.ScaleWidth - frmPrint.TextWidth("Página " & m_NumPag)
    
    frmPrint.Print "Página " & m_NumPag

End Sub


Private Sub Form_Load()

    Dim cnn                     As ADODB.Connection
    Dim sServidor               As String
    Dim sDataBase               As String
    Dim strUsuario              As String
    Dim strSenha                As String
    Dim sSql                    As String
    Dim rst                     As ADODB.Recordset
    Dim sDataProcessamento      As String
    Dim ndNode                  As Node
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    ''''''''''''''''''''''''''
    'Abre conexao com o banco'
    ''''''''''''''''''''''''''
    With cnn
        sServidor = PegarOpcaoINI("Conexao", "Servidor", "MDI_NT1")
        sDataBase = PegarOpcaoINI("Conexao", "DataBaseDestino", "Analisa")
        strUsuario = PegarOpcaoINI("Conexao", "Usuario", App.Path & "\MDI_Conexao.ini")
        strSenha = PegarOpcaoINI("Conexao", "Senha", App.Path & "\MDI_Conexao.ini")
        
        .Provider = "SQLOLEDB"
        .ConnectionTimeout = 5
        .CursorLocation = adUseClient
        .ConnectionString = "Server=" & sServidor & ";uid=" & strUsuario & ";pwd=" & strSenha & ";database=" & sDataBase
        .Open
    End With
    
           sSql = "SELECT CAP.DataProcessamento, EST.Nome_Estacao, EST.ID_Estacao"
    sSql = sSql & "  FROM Captura CAP, Estacao EST"
    sSql = sSql & " WHERE CAP.ID_Estacao = EST.ID_Estacao"
    
    rst.Open sSql, cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rst.EOF() Then
        Do While Not rst.EOF()
            If rst!DataProcessamento = Val(sDataProcessamento) Then
            
                Set ndNode = TreeView1.Nodes.Add("KEY_" & rst.AbsolutePosition, tvwChild, "KEY_" & rst!DataProcessamento & Trim(rst!Nome_Estacao), Trim(rst!Nome_Estacao), "COMPUTER")
                ndNode.Tag = rst!ID_Estacao
                rst.MoveNext
            Else
                sDataProcessamento = rst!DataProcessamento
                Set ndNode = TreeView1.Nodes.Add(, , "KEY_" & rst.AbsolutePosition, Format(Format(rst!DataProcessamento, "0000-00-00"), "dd-mm-yyyy"), "FLD_CLOSE", "FLD_OPEN")
            End If
        Loop
    End If
    
    rst.Close
    
    
    
    '''''''''''''''''''''''''''
    'Fecha conexao com o banco'
    '''''''''''''''''''''''''''
    cnn.Close

End Sub


