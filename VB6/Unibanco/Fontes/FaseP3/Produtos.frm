VERSION 5.00
Begin VB.Form Produtos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Produtos"
   ClientHeight    =   6456
   ClientLeft      =   1764
   ClientTop       =   1320
   ClientWidth     =   7044
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6456
   ScaleWidth      =   7044
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Selecionar"
      Height          =   336
      Left            =   2124
      Picture         =   "Produtos.frx":0000
      TabIndex        =   8
      Top             =   6012
      Width           =   1104
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   336
      Left            =   3804
      Picture         =   "Produtos.frx":030A
      TabIndex        =   7
      Top             =   6012
      Width           =   1104
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4668
      Left            =   36
      TabIndex        =   5
      Top             =   1200
      Width           =   6912
      Begin VB.ListBox LstProd 
         Height          =   4272
         Left            =   144
         TabIndex        =   6
         Top             =   228
         Width           =   6672
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar por :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1152
      Left            =   48
      TabIndex        =   0
      Top             =   24
      Width           =   6900
      Begin VB.CommandButton CmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   336
         Left            =   5676
         Picture         =   "Produtos.frx":0614
         TabIndex        =   9
         Top             =   612
         Width           =   1104
      End
      Begin VB.ComboBox CboSegmento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2904
      End
      Begin VB.TextBox TxtDescricao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Width           =   2292
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Segmento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   192
         TabIndex        =   4
         Top             =   348
         Width           =   876
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   3252
         TabIndex        =   3
         Top             =   348
         Width           =   900
      End
   End
End
Attribute VB_Name = "Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Veriáveis de Trabalho
Public Selecionou As Boolean
Public CodigoSelecionado As Integer

'Declaração de Variáveis do RDO
Private qryGetCONAX As rdoQuery
Sub AjustesIniciais()

  Dim sSql As String

  'Preencher o Combo com os tipos de segmentos
  CboSegmento.AddItem "1 - Prefeituras"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 1

  CboSegmento.AddItem "2 - Saneamento"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 2

  CboSegmento.AddItem "3 - Energia Elétrica e Gás"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 3

  CboSegmento.AddItem "4 - Telecomunicações"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 4

  CboSegmento.AddItem "5 - Secretarias Estaduais"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 5

  CboSegmento.AddItem "6 - Carnês e Assemelhados"
  CboSegmento.ItemData(CboSegmento.NewIndex) = 6

  'Preencher o List com todos os Produtos da tabela
  sSql = "0,"                                                       'Código do Produto
  sSql = sSql & "0,"                                                'Código do Segmento
  sSql = sSql & "'',"                                               'Descrição do Produto
  sSql = sSql & Format(CStr(Geral.capa.AgOrig), "0000") & ","       'Agencia de Coleta
  sSql = sSql & "5"                                                 'Tipo de Consulta

  Call PesquisaCONAX(sSql)

  'Posicionar o foco no campo 'DESCRIÇÃO'
  TxtDescricao.SetFocus
End Sub
Sub PesquisaCONAX(ByVal sSql As String)

  On Error GoTo ERRO_PESQUISACONAX

  Dim RsCONAX As rdoResultset

  'Pesquisar na tabela CONAX e preencher a lista
  Set qryGetCONAX = Geral.Banco.CreateQuery("", "{call GetCONAX (" & sSql & ")}")
  Set RsCONAX = qryGetCONAX.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsCONAX.EOF Then
    LstProd.Visible = False
    Do Until RsCONAX.EOF
      LstProd.AddItem RsCONAX!Descricao
      LstProd.ItemData(LstProd.NewIndex) = RsCONAX!Codigo

      RsCONAX.MoveNext
      DoEvents
    Loop
    LstProd.Visible = True
    LstProd.SetFocus
  Else
    MsgBox "Nenhum documento encontrado para este argumento.", vbInformation + vbOKOnly, App.Title

    TxtDescricao.SelStart = 0
    TxtDescricao.SelLength = Len(TxtDescricao.Text)
    TxtDescricao.SetFocus
  End If

  Exit Sub

ERRO_PESQUISACONAX:
  Select Case TratamentoErro("Erro ao Ler Produtos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Sub CboSegmento_Click()

  TxtDescricao.Text = ""
  If CboSegmento.ListIndex <> -1 Then
    Call CmdPesquisar_Click
  End If
End Sub
Private Sub CboSegmento_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call CmdPesquisar_Click
  End If
End Sub
Private Sub cmdConfirmar_Click()

  'Selecionar o Item Atual e retornar para a tela chamadora
  If LstProd.ListIndex <> -1 Then
    CodigoSelecionado = LstProd.ItemData(LstProd.ListIndex)
    Selecionou = True
    Me.Hide
  End If

  Exit Sub

ERRO_CONFIRMACONAX:
  Select Case TratamentoErro("Erro ao Selecionar Produto.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Private Sub CmdPesquisar_Click()

  On Error GoTo ERRO_PESQUISAR

  Dim sSql As String

  'Verificar se o usuário selecionou um segmento
  If CboSegmento.ListIndex <> -1 Then
    sSql = "0,"                                                       'Código do Produto
    sSql = sSql & CboSegmento.ItemData(CboSegmento.ListIndex) & ","   'Código do Segmento
    sSql = sSql & "'',"                                               'Descrição do Produto
'    sSql = sSql & Geral.AgenciaCentral & ","                          'Agencia Central
    sSql = sSql & Format(CStr(Geral.capa.AgOrig), "0000") & ","       'Agencia de Coleta
    sSql = sSql & "2"                                                 'Tipo de Consulta
  ElseIf Len(Trim(TxtDescricao.Text)) <> 0 Then
    sSql = "0,"                                                       'Código do Produto
    sSql = sSql & "0,"                                                'Código do Segmento
    sSql = sSql & "'" & TxtDescricao.Text & "',"                      'Descrição do Produto
'    sSql = sSql & Geral.AgenciaCentral & ","                          'Agencia Central
    sSql = sSql & Format(CStr(Geral.capa.AgOrig), "0000") & ","       'Agencia de Coleta
    sSql = sSql & "3"                                                 'Tipo de Consulta
  Else
    MsgBox "Informe um Critério de Seleção.", vbInformation, App.Title
    LstProd.Clear
    Exit Sub
  End If

  LstProd.Clear

  Call PesquisaCONAX(sSql)

  Exit Sub

ERRO_PESQUISAR:
  Select Case TratamentoErro("Erro ao Ler Produtos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Sub cmdSair_Click()

  Selecionou = False
  Me.Hide
End Sub
Private Sub Form_Activate()

  Call AjustesIniciais

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  'Fachar as Conexões
  Set qryGetCONAX = Nothing
End Sub

Private Sub LstProd_DblClick()

  Call cmdConfirmar_Click
End Sub
Private Sub LstProd_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    'Selecionar o Item Atual e retornar para a tela Chamadora
    Call cmdConfirmar_Click
  End If
End Sub


Private Sub TxtDescricao_GotFocus()

  TxtDescricao.SelStart = 0
  TxtDescricao.SelLength = Len(TxtDescricao.Text)
End Sub
Private Sub TxtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)

  CboSegmento.ListIndex = -1
  LstProd.Clear
End Sub
Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call CmdPesquisar_Click
  End If
End Sub
