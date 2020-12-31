VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form Titulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Títulos de Outros Bancos sem Código de Barras"
   ClientHeight    =   1896
   ClientLeft      =   1272
   ClientTop       =   1284
   ClientWidth     =   7932
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1896
   ScaleWidth      =   7932
   Begin VB.Frame Frame4 
      Height          =   1848
      Left            =   48
      TabIndex        =   8
      Top             =   -12
      Width           =   7848
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   6024
         Picture         =   "Titulo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   252
         Width           =   804
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   696
         Left            =   6840
         Picture         =   "Titulo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   252
         Width           =   816
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   5208
         Picture         =   "Titulo.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   252
         Width           =   804
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   696
         Left            =   4392
         Picture         =   "Titulo.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   252
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   3576
         Picture         =   "Titulo.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   252
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   2760
         Picture         =   "Titulo.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   252
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   1944
         Picture         =   "Titulo.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   252
         Width           =   804
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
         Height          =   372
         Left            =   2472
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1212
         Width           =   480
      End
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   372
         Left            =   4008
         TabIndex        =   1
         Top             =   1212
         Width           =   2160
         _Version        =   65537
         _ExtentX        =   3810
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
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
         Left            =   1764
         TabIndex        =   11
         Top             =   1284
         Width           =   576
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Títulos"
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
         Left            =   744
         TabIndex        =   10
         Top             =   480
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   216
         Picture         =   "Titulo.frx":1546
         Top             =   360
         Width           =   384
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         Left            =   3396
         TabIndex        =   9
         Top             =   1284
         Width           =   468
      End
   End
End
Attribute VB_Name = "Titulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Veriáveis de Trabalho
Public Alterou As Boolean
Public mForm As Form

'Declaração de veriável RDOQUERY
Private qryAtualizaTitulo As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryGetTitulo As rdoQuery
Sub AjustesIniciais()

  'Setar os objetos RDOQUERY
  Set qryAtualizaTitulo = Geral.Banco.CreateQuery("", "{call AtualizaTitulo (?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
End Sub
Sub PesquisaTitulo()

  On Error GoTo ERRO_PESQUISATITULO

  Dim sSql As String
  Dim RsTitulo As rdoResultset

  'Pesquisar o Titulo Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetTitulo = Geral.Banco.CreateQuery("", "{call GetTitulo (" & sSql & ")}")

  Set RsTitulo = qryGetTitulo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsTitulo.EOF Then
    'Encontrou o Titulo -> Preencher os campos
    txtBanco.Text = RsTitulo!Banco
    TxtValor.Text = RsTitulo!Valor * 100
  End If

  'Posicionando o foco no campo de valor
  txtBanco.SetFocus
  Call txtBanco_Change

  Exit Sub

ERRO_PESQUISATITULO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados do Depósito.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function SalvaTitulo() As Boolean

On Error GoTo ERRO_SALVATITULO
  
  Dim strEncripta   As String
  
  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Validar o Código do Banco
    If Not ValidaCodigoBanco(txtBanco.Text) Then
      MsgBox "Código de Banco não participante do Sistema de Compensação.", vbInformation, App.Title
      txtBanco.SetFocus
      Exit Function
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 12 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(12, CStr(Val(txtBanco.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVATITULO

    'Atualizar / Inserir Titulo
    With qryAtualizaTitulo
      .rdoParameters(0) = Geral.DataProcessamento                       'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto                       'IdDocto
      .rdoParameters(2) = Val(txtBanco.Text)                            'Banco
      .rdoParameters(3) = Val(TxtValor.Text) / 100                      'Valor
      .rdoParameters(4) = ""                                            'Leitura
      .rdoParameters(5) = 12                                            'TipoDocto
      .rdoParameters(6) = strEncripta                                   'Autenticacao digital
      .Execute
    End With

    SalvaTitulo = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 12
  End If

  Exit Function

ERRO_SALVATITULO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do Título.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Private Sub cmdConfirmar_Click()

  If SalvaTitulo Then
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

  Call PesquisaTitulo

  Screen.MousePointer = vbDefault
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

  Set qryAtualizaTitulo = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryGetTitulo = Nothing
End Sub

Private Sub txtBanco_Change()

  If Len(Trim(txtBanco.Text)) = txtBanco.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtBanco_GotFocus()

  txtBanco.SelStart = 0
  txtBanco.SelLength = txtBanco.MaxLength
End Sub
Function CamposOK() As Boolean

  CamposOK = False

  'Banco
  If Len(Trim(txtBanco.Text)) = 0 Then
    MsgBox "Informe o Número do Banco.", vbInformation, App.Title
    txtBanco.SetFocus
    Exit Function
  End If

  'Valor
  If Len(Trim(TxtValor.Text)) = 0 Then
    MsgBox "Informe o Valor do Título.", vbInformation, App.Title
    TxtValor.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function

Private Sub txtBanco_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtBanco_Validate(Cancel As Boolean)

   If Len(Trim(txtBanco.Text)) > 0 And Val(txtBanco.Text) <> 0 Then
      'Validar o Código do Banco
      If Not ValidaCodigoBanco(txtBanco.Text) Then
         MsgBox "Código de Banco não participante do Sistema de Compensação.", vbInformation, App.Title
         Cancel = True
         txtBanco.SelStart = 0
         txtBanco.SelLength = Len(txtBanco.Text)
      End If
   End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
