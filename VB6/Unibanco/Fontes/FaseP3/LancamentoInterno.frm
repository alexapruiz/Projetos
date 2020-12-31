VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form LancamentoInterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Lançamento Interno"
   ClientHeight    =   2772
   ClientLeft      =   1032
   ClientTop       =   4896
   ClientWidth     =   10152
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2772
   ScaleWidth      =   10152
   Begin VB.Frame Frame1 
      Caption         =   "Confirmação do Controle do Banco"
      Height          =   828
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   9948
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   336
         Left            =   7104
         TabIndex        =   20
         Top             =   312
         Width           =   972
      End
      Begin VB.TextBox txtBanco1 
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
         Left            =   4488
         MaxLength       =   21
         TabIndex        =   18
         Top             =   288
         Width           =   2448
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Controle do Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   2424
         TabIndex        =   19
         Top             =   384
         Width           =   1596
      End
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   800
      Left            =   8352
      Picture         =   "LancamentoInterno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   800
      Left            =   9216
      Picture         =   "LancamentoInterno.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   800
      Left            =   4872
      Picture         =   "LancamentoInterno.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   800
      Left            =   4008
      Picture         =   "LancamentoInterno.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5736
      Picture         =   "LancamentoInterno.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter Cor"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   6600
      Picture         =   "LancamentoInterno.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Frente/Verso"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7464
      Picture         =   "LancamentoInterno.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   96
      Width           =   876
   End
   Begin VB.Frame fraCartao 
      Height          =   1788
      Left            =   96
      TabIndex        =   15
      Top             =   960
      Width           =   9948
      Begin CURRENCYEDITLib.CurrencyEdit txtValorOperacao 
         Height          =   360
         Left            =   6528
         TabIndex        =   3
         Top             =   1248
         Width           =   2580
         _Version        =   65537
         _ExtentX        =   4551
         _ExtentY        =   635
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   11
         BackColor       =   -2147483643
      End
      Begin DATEEDITLib.DateEdit dtData_Validade 
         Height          =   372
         Left            =   2040
         TabIndex        =   0
         Top             =   768
         Width           =   1572
         _Version        =   65537
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtBanco 
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
         Left            =   6528
         MaxLength       =   21
         TabIndex        =   2
         Top             =   768
         Width           =   2592
      End
      Begin VB.TextBox txtEvento 
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
         Left            =   6528
         MaxLength       =   4
         TabIndex        =   1
         Top             =   288
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Geração"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   1416
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Controle do Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   4464
         TabIndex        =   5
         Top             =   864
         Width           =   1596
      End
      Begin VB.Label lblValorOperacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Operação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   4464
         TabIndex        =   6
         Top             =   1344
         Width           =   1560
      End
      Begin VB.Label lblEvento 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código do Evento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   4464
         TabIndex        =   4
         Top             =   384
         Width           =   1536
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de Lançamento Interno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   768
      TabIndex        =   14
      Top             =   360
      Width           =   2748
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   240
      Picture         =   "LancamentoInterno.frx":1546
      Top             =   240
      Width           =   384
   End
End
Attribute VB_Name = "LancamentoInterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variavel de retorno informando se Cancelou ou Alterou
Public Alterou                              As Boolean
Dim sPosicaoErro                            As String
Dim dcontrolebco                            As String
Private mForm                               As Form
Private Type tpModulo
    qryInserirLanctoInterno                 As rdoQuery
    qryGetLancamentoInterno                 As rdoQuery
    qryRemoveTipoDocumento                  As rdoQuery
    qryValidaEventoLancto                   As rdoQuery
End Type

'Declaração da DLL calculo do NSU e Calculo do SDV
Private Declare Function QXCalNsu Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal Ret As String) As Integer
Private Declare Function QXGetSDV Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal CIF As String, ByVal Ret As String) As Integer

Private ControleBanco                       As String
Private Modulo                              As tpModulo
Private DuplaDigitacao                      As Boolean

Private Function VerificarTudo() As Boolean

    Dim strEvento As String

    VerificarTudo = False

    'Verifica data de geração
    If Len(dtData_Validade.Text) = 0 Then
        MsgBox "O preenchimento da Data de Geração é obrigatório !", vbInformation + vbOKOnly, App.Title
        dtData_Validade.SetFocus
        Exit Function
    End If

    'Verifica dígito do campo Evento
    If Len(txtEvento.Text) = 0 Or Val(txtEvento.Text) = 0 Then
        MsgBox "O preenchimento do código de Evento é obrigatório !", vbInformation + vbOKOnly, App.Title
        txtEvento.SetFocus
        Exit Function
    Else
        If ValidaEventoLancto(txtEvento.Text) = False Then
            MsgBox "Código de Evento inválido.", vbInformation + vbOKOnly, App.Title
            txtEvento.SelStart = 0
            txtEvento.SelLength = Len(txtEvento.Text)
            txtEvento.SetFocus
            Exit Function
        End If
    End If

    strEvento = CStr(Val(txtEvento.Text))
    txtEvento = Format(txtEvento.Text, "0000")

    'Valida o campo valor
    If Val(Desformata_Valor(txtValorOperacao.Text)) = 0 Then
        MsgBox "Digite o Valor da operação.", vbInformation + vbOKOnly, App.Title
        txtValorOperacao.SetFocus
        Exit Function
    End If

    'If UCase(dcontrolebco) <> 0 Then
    '    VerificarTudo = True
    '    Exit Function
    'End If
    
    'Primeira digitação
    If Len(Trim(txtBanco.Text)) = 0 And txtBanco1.Visible = False Then
        MsgBox "O preenchimento de Controle do Banco é obrigatório !", vbInformation + vbOKOnly, App.Title
        txtBanco.SetFocus
        Exit Function
    Else
        dcontrolebco = UCase(IIf(Len(Trim(txtBanco.Text)) > 0, txtBanco.Text, txtBanco1.Text))
        If Len(Trim(dcontrolebco)) > 14 Then
            If CalculaControleBanco_NovoNSU(dcontrolebco) = False Then
                MsgBox "Controle do Banco é inválido!", vbInformation + vbOKOnly, App.Title
                txtBanco.SelStart = 0
                txtBanco.SelLength = Len(txtBanco)
                txtBanco.SetFocus
                Exit Function
            End If
        Else
            If CalculaControleBanco(dcontrolebco) = False Then
                MsgBox "Controle do Banco é inválido!", vbInformation + vbOKOnly, App.Title
                txtBanco.SelStart = 0
                txtBanco.SelLength = Len(txtBanco)
                txtBanco.SetFocus
                Exit Function
            End If
        End If
    End If
    
    VerificarTudo = True

End Function
Private Sub cmdCancelar_Click()
    txtBanco.Text = dcontrolebco
    Me.Height = 3084
    dcontrolebco = "0"
    DuplaDigitacao = False
End Sub

Private Sub cmdConfirmar_Click()

    Dim bDuplicidade    As Boolean
    Dim dValor          As Double
    Dim strEncripta     As String
    
    On Error GoTo Err_SalvaDados:

    If VerificarTudo = False Then Exit Sub
    
    If DuplaDigitacao = False Then
       Me.Height = 3936
       txtBanco.Text = ""
       txtBanco1.SetFocus
       Exit Sub
    Else
       Me.ScaleHeight = 3084
       txtBanco.Text = dcontrolebco
       dcontrolebco = "0"
    End If
    
    sPosicaoErro = "InsLanctoInterno"
    dValor = Val(txtValorOperacao.Text) / 100

    ''''''''''''''''''
    'Inicia Transação'
    ''''''''''''''''''
    Geral.Banco.BeginTrans
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Verificar se o Documento pertence à outro Tipo'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If Geral.Documento.TipoDocto <> 41 And Geral.Documento.TipoDocto <> 0 Then
        With Modulo.qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
        End With
    End If
    
    Geral.Documento.TipoDocto = etpdocLancamentoInterno
    
    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(41, CStr(Val(Right(txtBanco.Text, 14))))
    If strEncripta = "" Then GoTo Exit_SalvaDados
    
    With Modulo.qryInserirLanctoInterno
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto
        .rdoParameters(3) = Mid(dtData_Validade.Text, 5, 4) & Mid(dtData_Validade.Text, 3, 2) & Mid(dtData_Validade.Text, 1, 2)
        .rdoParameters(4) = txtEvento.Text
        .rdoParameters(5) = txtBanco.Text
        .rdoParameters(6) = dValor
        .rdoParameters(7) = Geral.Documento.TipoDocto
        .rdoParameters(8) = strEncripta                             'Autenticacao digital
        .Execute
        
        If .rdoParameters(0).Value = 2 Then Geral.Documento.Status = "D"
    
        If .rdoParameters(0).Value = 0 Then GoTo Exit_SalvaDados
        
    'If .rdoParameters(0).Value <> 0 Then GoTo Exit_SalvaDados

    End With
    
                                        '''''''''''''''''''''''''''''''''''''''''''''
    Geral.Documento.Leitura = ""        'Não há necessidade de guardar campo leitura'
                                        '''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''
    'Atualiza tabela Documento'
    '''''''''''''''''''''''''''
    bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
    
    If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , etpdocLancamentoInterno, , , , , dValor) Then
        GoTo Exit_SalvaDados
    End If
     'Atualiza (bDuplicidade) fora da função (G_AtualizaCamposDocumento) devido ao campo
     'leitura (CHAVE) não conter o numero de CTRL do BANCO do LCto Interno
     If Geral.Documento.Status = "D" Then bDuplicidade = True
     
     ''''''''''''''''''''
     'Finaliza Transação'
     ''''''''''''''''''''
     Geral.Banco.CommitTrans
     
     Alterou = True
     Me.Hide
    
Exit Sub

Exit_SalvaDados:
    Alterou = False
    Geral.Banco.RollbackTrans
    MsgBox "Não foi possível Incluir/atualiza dados de Lancamento Interno.", vbInformation + vbOKOnly, App.Title
    ' cmdSair_Click
    Exit Sub
    
Err_SalvaDados:
    Alterou = False
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível inserir/atualizar o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbCancel
            Alterou = False
            Me.Hide
        Case vbRetry
    End Select

End Sub
Private Sub cmdFrenteVerso_Click()
    mForm.cmdFrenteVerso_Click
End Sub
Private Sub cmdInverteCor_Click()
    mForm.cmdInverteCor_Click
End Sub
Private Sub cmdOk_Click()
    
    If Len(Trim(txtBanco1.Text)) = 0 Then
        MsgBox "Entre com a confirmação do Controle do Banco.", vbInformation + vbOKOnly, App.Title
        txtBanco1.SetFocus
        Exit Sub
    End If

    If UCase(dcontrolebco) = UCase(txtBanco1.Text) Then
        DuplaDigitacao = True
        cmdConfirmar_Click
    Else
        MsgBox "Controle do Banco inválido", vbInformation + vbOKOnly, App.Title
        txtBanco1_GotFocus
    End If
End Sub

Private Sub cmdRotacao_Click()
    mForm.cmdRotacao_Click
End Sub
Private Sub CmdSair_Click()
    Alterou = False
    Me.Hide
End Sub
Private Sub cmdZoomMais_Click()
    mForm.cmdZoomMais_Click
End Sub
Private Sub cmdZoomMenos_Click()
    mForm.cmdZoomMenos_Click
End Sub
Private Sub dtData_Validade_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        If dtData_Validade.Locked = False Then
            dtData_Validade.Text = Mid(Geral.DataProcessamento, 7, 2) & _
                                   Mid(Geral.DataProcessamento, 5, 2) & _
                                   Mid(Geral.DataProcessamento, 1, 4)
        End If
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Form_Activate()
'    Dim iErroData  As Integer
    
    'Verifica se feriado na agência
    If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, "", False) Then
        CmdSair_Click
        Exit Sub
    End If
    
'    iErroData = ValidaAgencia(Geral.Capa.AgOrig, "", False)
'    If iErroData <> 0 Then
'        Select Case iErroData
'            Case 2 'Feriado
'                MsgBox "A Agência de Origem está em Feriado.", vbInformation + vbOKOnly, App.Title
'            Case 3 'Agência Fechada
'                MsgBox "A Agência de Origem está Fechada.", vbInformation + vbOKOnly, App.Title
'        End Select
'        If iErroData = 2 Or iErroData = 3 Then
'            CmdSair_Click
'            Exit Sub
'        End If
'    End If
    
    'Ler dados da tabela Cartão Avulso, pode ou não existir registro
    If Not LerDados() Then CmdSair_Click: Exit Sub
    
    dtData_Validade.SetFocus
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        txtBanco.Text = dcontrolebco
        Me.Height = 3084
        dcontrolebco = 0
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      cmdZoomMais_Click
    Case vbKeySubtract
      cmdZoomMenos_Click
    Case vbKeyF10
      cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      cmdRotacao_Click
    Case vbKeyF11
      cmdFrenteVerso_Click
    Case vbKeyMultiply
        Call cmdConfirmar_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        mForm.Form_KeyUp KeyCode, Shift
    Case 27
        CmdSair_Click
  End Select
  
End Sub
Private Sub Form_Load()

    Alterou = False
    
    '* Valor Default para Dupla Digitação
    DuplaDigitacao = False
    
    With Modulo
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Cria a query para a gravação dos dados do cartão'
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        Set .qryInserirLanctoInterno = Geral.Banco.CreateQuery("", "{? = call InserirLancamentoInterno (?,?,?,?,?,?,?,?)}")
            'Parametros: (1)-Data Processamento (2)-IdDocto (3)-Numero Cartao (4)-Desp.Reais (5)-Desp.Dolar (6)-Ant.Saque  (7)-Valor
            .qryInserirLanctoInterno.rdoParameters(0).Direction = rdParamReturnValue
            
        '''''''''''''''''''''''''''''''''''''''
        'Cria a query para Ler dados do cartão'
        '''''''''''''''''''''''''''''''''''''''
        Set .qryGetLancamentoInterno = Geral.Banco.CreateQuery("", "{? = call GetLancamentoInterno (?,?)}")
            'Parametros (1)-Data Processamento (2)-IdDocto
            .qryGetLancamentoInterno.rdoParameters(0).Direction = rdParamReturnValue
                
        Set .qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
        ''''''''''''
        'Cria query'
        ''''''''''''
        Set .qryValidaEventoLancto = Geral.Banco.CreateQuery("", "{? = call GetValidaEventoLancto (?)}")
        
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Fecha Recordset
    With Modulo
        .qryGetLancamentoInterno.Close
        .qryInserirLanctoInterno.Close
        .qryRemoveTipoDocumento.Close
    End With
End Sub

Private Sub txtBanco_GotFocus()
    SelecionarTexto txtBanco
End Sub
Private Sub txtBanco_KeyPress(KeyAscii As Integer)

    Dim sDif As Integer

    If KeyAscii = vbKeyReturn Then

        If Len(Trim(txtBanco.Text)) > 0 Then
            If Len(Trim(txtBanco.Text)) > 15 Then
                txtBanco.Text = Left(Trim(txtBanco.Text) & String(21, "0"), 21)
                sDif = 0
            Else
                txtBanco.Text = Left(Trim(txtBanco.Text) & String(14, "0"), 14)
                sDif = 7
            End If

            'Validação do Controle do Banco
            If Mid(txtBanco, 15 - sDif, 4) <> Format(Geral.Capa.AgOrig, "0000") Then
                MsgBox "Controle de Banco não confere com Agência Origem.", vbInformation + vbOKOnly, App.Title
                txtBanco.SelStart = 0
                txtBanco.SelLength = Len(txtBanco.Text)
                txtBanco.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        End If
    Else

        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or _
            (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
        Else
            If (KeyAscii >= 97 And KeyAscii <= 122) Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        End If
    End If
End Sub
Private Sub txtBanco1_GotFocus()
    SelecionarTexto txtBanco1
End Sub
Private Sub txtBanco1_KeyPress(KeyAscii As Integer)
    
    'InibirTeclaAlfa KeyAscii

    If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
    Else
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or (KeyAscii = vbKeyBack) Or _
            (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
        Else
            If (KeyAscii >= 97 And KeyAscii <= 122) Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        End If
    
    End If
        
End Sub

Private Sub txtEvento_GotFocus()
    SelecionarTexto txtEvento
End Sub
Private Sub txtEvento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        InibirTeclaAlfa KeyAscii
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEvento_LostFocus()
    If Len(Trim(txtEvento.Text)) > 0 Then
        'Valida Código do Evento
        If ValidaEventoLancto(txtEvento.Text) = False Then
            MsgBox "Código do Evento inválido.", vbInformation + vbOKOnly, App.Title
            txtEvento.Text = ""
            txtEvento.SetFocus
        Else
            txtBanco.SetFocus
        End If
    End If
End Sub
Public Sub SetParent(ByRef aForm As Form)
  Set mForm = aForm
End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)
  Me.Left = iLeft
  Me.Top = iTop
End Sub
Private Function LerDados() As Boolean

Dim rstModulo As rdoResultset

On Error GoTo Err_LerDados
    
    LerDados = False
    
    With Modulo.qryGetLancamentoInterno
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto
        Set rstModulo = .OpenResultset(rdOpenStatic)
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível ler dados de Lancamento Interno.", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If
        
        If Not rstModulo.EOF() Then
            Geral.Documento.Leitura = ""
            'Atualiza dados de entrada
            dtData_Validade.Text = Mid(rstModulo!DataGeracao, 7, 2) & Mid(rstModulo!DataGeracao, 5, 2) & Mid(rstModulo!DataGeracao, 1, 4)
            txtEvento.Text = rstModulo!Evento
            txtBanco.Text = Format(rstModulo!ControleBanco, String(14, "0"))
            txtValorOperacao.Text = rstModulo!Valor * 100
        Else
            Geral.Documento.Leitura = ""
            'Atualiza dados de entrada
            txtEvento.Text = ""
            txtBanco.Text = ""
            txtValorOperacao.Text = 0
        End If
    End With
    
    LerDados = True

Exit_LerDados:
    Set rstModulo = Nothing
    Exit Function
    
Err_LerDados:

    Select Case TratamentoErro("Não foi possível ler dados referentes a Lancamento Interno !", Err, rdoErrors, False)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    GoTo Exit_LerDados

End Function
Function ValidaEventoLancto(Codigo As Long) As Boolean

    'Valida Código do Evento do Lançamento Interno de acordo com tabela
    Dim RsValidaEvento As rdoResultset

    With Modulo.qryValidaEventoLancto
        .rdoParameters(1).Value = Codigo
        Set RsValidaEvento = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If RsValidaEvento.EOF Then
        ValidaEventoLancto = False
    Else
        ValidaEventoLancto = True
    End If
End Function
Private Sub txtValorOperacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        InibirTeclaAlfa KeyAscii
        Call cmdConfirmar_Click
    End If
End Sub
Private Function CalculaControleBanco(NumControleBanco As String) As Boolean
   
    Dim soma As Integer
    Dim resto As Integer
    Dim digito_z As Integer
    Dim digito_x As Integer
    Dim p As Integer
    Dim peso As Integer
    Dim calcula_x As String
    Dim calcula_z As String
    Dim base As String * 7
    
    '---- Calculo do Controle do Banco ----'
    ' 1 - Base NSU
    ' 2 - 0 (default)
    ' 3 - Agência Origem
    ' 4 - Numero do Terminal
    ' 5 - Data de Geração

    calcula_x = Format(Mid(NumControleBanco, 1, 4), "0000")
    calcula_x = calcula_x & "0"
    calcula_x = calcula_x & Format(Mid(NumControleBanco, 8, 4), "0000")
    calcula_x = calcula_x & Format(Mid(NumControleBanco, 12, 3), "000")
    calcula_x = calcula_x & Mid(dtData_Validade.Text, 1, 2) & Mid(dtData_Validade.Text, 3, 2) & Mid(dtData_Validade.Text, 7, 2)
       
    soma = 0
    resto = 0
    digito_x = 0        'calculado pelo módulo 11
    
    '*************************************************************
    'número base: (18)                     SSSS0aaaatttddmmaa
    '                                      xxxxxxxxXXXXXXXXXX
    'multiplica da esquerda para direita:  234567892345678923
    '*************************************************************
    
    peso = 2    'começa multiplicar da direita para esquerda
    p = 18
   
    Do
        '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
        soma = soma + Val(Mid(calcula_x, p, 1)) * peso
        p = p - 1            'ponteiro
        peso = peso + 1      'peso
        If (peso = 10) Then
            peso = 2
        End If
        If (p = 0) Then
            Exit Do
        End If
    Loop
   
    resto = soma Mod 11        'resto da divisão
    digito_x = 11 - resto     'digito verificador
   
    '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
    If (digito_x = 11) Or (digito_x = 10) Then
        digito_x = 0
    End If
      
    '---- Proximo calculo ---'
    ' 1 - Base NSU
    ' 2 - Digito_x
    ' 3 - 0
    ' 4 - Agencia Central
    ' 5 - Terminal
    ' 6 - Data de Geração
       
    calcula_z = Format(Mid(NumControleBanco, 1, 4), "0000")
    calcula_z = calcula_z & Format((digito_x), "0")
    calcula_z = calcula_z & "0"
    calcula_z = calcula_z & Format(Mid(NumControleBanco, 8, 4), "0000")
    calcula_z = calcula_z & Format(Mid(NumControleBanco, 12, 3), "000")
    calcula_z = calcula_z & Mid(dtData_Validade.Text, 1, 2) & Mid(dtData_Validade.Text, 3, 2) & Mid(dtData_Validade.Text, 7, 2)
      
    soma = 0
    resto = 0
    digito_z = 0        'calculado pelo módulo 11
      
    '*************************************************************
    'número base: (19)                     SSSSX0aaaatttddmmyy
    '                                      xxxxxxxxXXXXXXXXXXx
    'multiplica da esquerda para direita:  2345678923456789234
    '*************************************************************
   
    peso = 2    'começa multiplicar da direita para esquerda
    p = 19
   
    Do
        '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
        '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
        soma = soma + Val(Mid(calcula_z, p, 1)) * peso
        p = p - 1            'ponteiro
        peso = peso + 1      'peso
        If (peso = 10) Then
            peso = 2
        End If
        If (p = 0) Then
            Exit Do
        End If
    Loop
   
    resto = soma Mod 11        'resto da divisão
    digito_z = 11 - resto     'digito verificador
   
    '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
    If (digito_z = 11) Or (digito_z = 10) Then
        digito_z = 0
    End If
      
     If Mid(NumControleBanco, 5, 2) = Format(digito_x, "0") & Format(digito_z, "0") Then
        CalculaControleBanco = True
     Else
        CalculaControleBanco = False
     End If
     
End Function
Private Function CalculaControleBanco_NovoNSU(ByVal ControleBanco As String) As Boolean

    Dim Ret         As Integer
    Dim NSU         As String * 6
    Dim SDV         As String * 1
    Dim Data        As String
    Dim nroCaixa    As String
    Dim Agencia     As String
    Dim TipoAgencia As String
    Dim Valor       As String
    Dim Operador    As String

    'Preenche as variáveis  que serão usadas pela DLL para calcular o Novo NSU
    Data = Left(dtData_Validade.Text, 4) & Mid(dtData_Validade.Text, 7, 2)
    nroCaixa = Right(ControleBanco, 3)
    Agencia = Format(Geral.Capa.AgOrig, "0000")
    Valor = txtValorOperacao.Text
    Operador = Mid(ControleBanco, 2, 6)
    TipoAgencia = Mid(ControleBanco, 14, 1)
    NSU = Mid(ControleBanco, 8, 4)

    NSU = Format(NSU, "0000") & "00"
    'Chama a DLL para o cálculo do NSU
    Ret = QXCalNsu(Agencia, nroCaixa, Data, TipoAgencia, Valor, NSU)

    'Verifica se não houve erro e se os DVs estão OK
    If Ret <> 0 Or NSU <> Mid(ControleBanco, 8, 6) Then Exit Function

    'Chama a DLL para o cálculo do Super DV
    Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)

    'Verifica se não houve erro e se os DVs estão OK
    If Ret <> 0 Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
    
    CalculaControleBanco_NovoNSU = True
End Function

