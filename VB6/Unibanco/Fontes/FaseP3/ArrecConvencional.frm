VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form ArrecConvencional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrecadação Convencional"
   ClientHeight    =   1968
   ClientLeft      =   1320
   ClientTop       =   1320
   ClientWidth     =   9288
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1968
   ScaleWidth      =   9288
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   12
      TabIndex        =   10
      Top             =   24
      Width           =   9240
      Begin VB.CommandButton CmdListaProduto 
         Caption         =   "Lista de Produtos"
         Height          =   660
         Left            =   7584
         Picture         =   "ArrecConvencional.frx":0000
         TabIndex        =   15
         Top             =   204
         Width           =   780
      End
      Begin VB.TextBox txtRequisicao 
         Enabled         =   0   'False
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
         Height          =   372
         Left            =   3528
         MaxLength       =   7
         TabIndex        =   1
         Top             =   1356
         Width           =   1308
      End
      Begin VB.TextBox txtProduto 
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
         Height          =   372
         Left            =   1188
         MaxLength       =   4
         TabIndex        =   0
         Top             =   1356
         Width           =   588
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   660
         Left            =   2832
         Picture         =   "ArrecConvencional.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   660
         Left            =   3624
         Picture         =   "ArrecConvencional.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   660
         Left            =   4416
         Picture         =   "ArrecConvencional.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   660
         Left            =   5208
         Picture         =   "ArrecConvencional.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   660
         Left            =   6000
         Picture         =   "ArrecConvencional.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   660
         Left            =   8376
         Picture         =   "ArrecConvencional.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   204
         Width           =   780
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   660
         Left            =   6792
         Picture         =   "ArrecConvencional.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   204
         Width           =   780
      End
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   372
         Left            =   6636
         TabIndex        =   2
         Top             =   1356
         Width           =   2076
         _Version        =   65537
         _ExtentX        =   3662
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
         MaxLength       =   11
         BackColor       =   -2147483643
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3552
         TabIndex        =   11
         Top             =   1080
         Width           =   996
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1128
         TabIndex        =   12
         Top             =   1044
         Width           =   696
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   168
         Picture         =   "ArrecConvencional.frx":1850
         Top             =   312
         Width           =   384
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Arrecadação Convencional"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   684
         TabIndex        =   13
         Top             =   432
         Width           =   1968
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6672
         TabIndex        =   14
         Top             =   1080
         Width           =   468
      End
   End
End
Attribute VB_Name = "ArrecConvencional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean

'Declaração das variáveis do RDO
Private qryRemoveTipoDocumento As rdoQuery
Private qryGetArrecConv As rdoQuery
Private qryAtualizaArrecConv As rdoQuery
Private qryGetCONAX As rdoQuery
Sub AjustesIniciais()

  'Setando as Variáveis do RDO
  Set qryAtualizaArrecConv = Geral.Banco.CreateQuery("", "{call AtualizaArrecConv (?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryGetCONAX = Geral.Banco.CreateQuery("", "{call GetCONAX (?,?,?,?,?)}")
End Sub
Function CamposOK() As Boolean

  CamposOK = False

  'Preenchimento do Código do Produto
  If Len(Trim(txtProduto.Text)) = 0 Then
    MsgBox "Informe o Código do Produto.", vbInformation, App.Title
    CamposOK = False
    txtProduto.SetFocus
    Exit Function
  End If

  'Requisicao
  If txtRequisicao.Enabled = True And Val(txtRequisicao.Text) = 0 Then
    MsgBox "Informe o código da Requisição.", vbInformation, App.Title
    CamposOK = False
    txtRequisicao.SetFocus
    Exit Function
  End If

  'Valor da Arrecadação
  If Len(Trim(TxtValor.Text)) = 0 Then
    MsgBox "Informe o Valor da Arrecadação.", vbInformation, App.Title
    CamposOK = False
    TxtValor.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Sub PesquisaArrecConv()

  On Error GoTo ERRO_PESQUISAARRECCONV

  Dim sSql As String
  Dim RsArrecConv As rdoResultset

  'Pesquisar o Documento Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetArrecConv = Geral.Banco.CreateQuery("", "{call GetArrecConv (" & sSql & ")}")

  Set RsArrecConv = qryGetArrecConv.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsArrecConv.EOF Then
    'Encontrou o Documento -> Preencher os campos
    txtProduto.Text = RsArrecConv!Produto
    If Val(RsArrecConv!Requisicao) <> 0 Then
      txtRequisicao.Text = RsArrecConv!Requisicao
    Else
      txtRequisicao.Text = ""
    End If

    TxtValor.Text = CCur(RsArrecConv!Valor) * 100

    'Posicionando o foco no campo de valor
    TxtValor.SetFocus
    TxtValor.SelStart = 0
    TxtValor.SelLength = Len(TxtValor.Text)
  Else
    'Posicionar o Foco no campo 'CÓDIGO DO PRODUTO'
    txtProduto.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PESQUISAARRECCONV:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados da Arrecadação Convencional.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function SalvaArrec() As Boolean

  On Error GoTo ERRO_SALVAARREC

  Dim sSql          As String
  Dim RsCONAX       As rdoResultset
  Dim strEncripta   As String
  
  SalvaArrec = False

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Validar Código do Produto (tabela : 'CONAX')
    sSql = txtProduto.Text & ","                                      'Código do Produto
    sSql = sSql & "0,"                                                'Código do Segmento
    sSql = sSql & "'',"                                               'Descrição do Produto
    sSql = sSql & Format(CStr(Geral.Capa.AgOrig), "0000") & ","       'Agencia de Coleta
    sSql = sSql & "1"                                                 'Tipo de Consulta

    Set qryGetCONAX = Geral.Banco.CreateQuery("", "{call Getconax (" & sSql & ")}")

    Set RsCONAX = qryGetCONAX.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    If RsCONAX.EOF Then
      MsgBox "O Código do Produto não é Válido.", vbInformation, App.Title
      txtProduto.SetFocus
      Exit Function
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 27 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If
    
    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(27, CStr(Val(txtProduto.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVAARREC

    'Atualizar / Inserir Arrecadação Convencional
    With qryAtualizaArrecConv
      .rdoParameters(0) = Geral.DataProcessamento                 'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto                 'IdDocto
      .rdoParameters(2) = Val(txtProduto.Text)                    'Código do Produto
      .rdoParameters(3) = Val(txtRequisicao.Text)                 'Requisicao
      .rdoParameters(4) = Val(TxtValor.Text) / 100                'Valor
      .rdoParameters(5) = ""                                      'Leitura
      .rdoParameters(6) = 27                                      'TipoDocto
      .rdoParameters(7) = strEncripta                             'Autenticacao digital
      .Execute
    End With

    SalvaArrec = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.TipoDocto = 27
    Geral.Documento.Leitura = ""
  End If

  Exit Function

ERRO_SALVAARREC:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Arrecadação Convencional.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdConfirmar_Click()

  If SalvaArrec Then
    Alterou = True
    Me.Hide
  End If
End Sub

Private Sub cmdFrenteVerso_Click()

  mForm.cmdFrenteVerso_Click
End Sub

Private Sub cmdInverteCor_Click()

  mForm.cmdInverteCor_Click
End Sub

Private Sub CmdListaProduto_Click()

  Produtos.Show vbModal, Me

  If Produtos.Selecionou Then
    DoEvents
    txtProduto.Text = Produtos.CodigoSelecionado

    Call VerificaProdutoRequisicao

    txtProduto.SelStart = 0
    txtProduto.SelLength = txtProduto.MaxLength
    txtProduto.SetFocus
  End If

  Unload Produtos
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


Private Sub Form_Activate()

  Call AjustesIniciais
  
  Call PesquisaArrecConv
End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub




Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      Call cmdZoomMais_Click
    Case vbKeySubtract
      Call cmdZoomMenos_Click
    Case vbKeyF10
      Call cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      Call cmdRotacao_Click
    Case vbKeyMultiply
      Call cmdConfirmar_Click
    Case vbKeyF11
      Call cmdFrenteVerso_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
      mForm.Form_KeyUp KeyCode, Shift
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Set qryAtualizaArrecConv = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryGetCONAX = Nothing
End Sub


Private Sub txtProduto_Change()
  'If Len(Trim(txtProduto.Text)) = txtProduto.MaxLength Then
  '  SendKeys "{TAB}"
  'End If
End Sub
Private Sub txtProduto_GotFocus()

  txtProduto.SelStart = 0
  txtProduto.SelLength = txtProduto.MaxLength
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If Len(Trim(txtProduto.Text)) = 0 Then
      txtRequisicao.Text = ""
      Call CmdListaProduto_Click
    Else
      SendKeys "{TAB}"
    End If
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtProduto_LostFocus()

    If Len(Trim(txtProduto)) = 0 Then Exit Sub
    '* Valida Código do Produto *'
    If Not IsNumeric(txtProduto) Then
        MsgBox "Código do Produto inválido, Redigite.", vbInformation, App.Title
        txtProduto.Text = ""
        txtProduto.SetFocus
    Else
        Call VerificaProdutoRequisicao
    End If


End Sub
Sub VerificaProdutoRequisicao()
  'Se Código do Produto = 3160 ou 3170 -> A Requisicao é Obrigatória
  If Val(txtProduto.Text) = 3160 Or Val(txtProduto.Text) = 3170 Or Val(txtProduto.Text) = 7028 Then
      txtRequisicao.Enabled = True
      txtRequisicao.SetFocus
  Else
      txtRequisicao.Enabled = False
      txtRequisicao.Text = ""
  End If
End Sub
Private Sub txtRequisicao_GotFocus()
  txtRequisicao.SelStart = 0
  txtRequisicao.SelLength = Len(txtRequisicao.Text)
End Sub
Private Sub txtRequisicao_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtRequisicao_LostFocus()
    '* Valida Código de Requisição *'
    If Len(Trim(txtRequisicao.Text)) = 0 Then Exit Sub
        If Not IsNumeric(txtRequisicao.Text) Then
            MsgBox "Numero de Requisição inválido, Redigite!", vbInformation, App.Title
            txtRequisicao.Text = ""
            txtRequisicao.SetFocus
            Exit Sub
        End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
