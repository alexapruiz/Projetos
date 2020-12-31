VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form ArrecEletronica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Arrecadações Eletrônicas"
   ClientHeight    =   1944
   ClientLeft      =   1284
   ClientTop       =   1308
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1944
   ScaleWidth      =   8880
   Begin VB.TextBox txtCodigo2 
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
      Left            =   1761
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1476
      Width           =   1548
   End
   Begin VB.TextBox txtCodigo3 
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
      Left            =   3402
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1476
      Width           =   1548
   End
   Begin VB.TextBox txtCodigo1 
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
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1476
      Width           =   1548
   End
   Begin VB.TextBox txtCodigo4 
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
      Left            =   5043
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1476
      Width           =   1548
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   7980
      Picture         =   "ArrecEletronica.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   72
      Width           =   816
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Frente/Verso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   6336
      Picture         =   "ArrecEletronica.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   72
      Width           =   816
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   5508
      Picture         =   "ArrecEletronica.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   72
      Width           =   816
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
      Height          =   696
      Left            =   4680
      Picture         =   "ArrecEletronica.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   72
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   3852
      Picture         =   "ArrecEletronica.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   72
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   3024
      Picture         =   "ArrecEletronica.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   72
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   7152
      Picture         =   "ArrecEletronica.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   72
      Width           =   816
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
      Height          =   372
      Left            =   6744
      TabIndex        =   4
      Top             =   1476
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
      Locked          =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Arrecadação Eletrônica"
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   720
      TabIndex        =   14
      Top             =   324
      Width           =   1764
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   204
      Picture         =   "ArrecEletronica.frx":1546
      Top             =   180
      Width           =   384
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   6804
      TabIndex        =   13
      Top             =   1212
      Width           =   492
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Linha Digitável do Código de Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   120
      TabIndex        =   12
      Top             =   1140
      Width           =   3396
   End
End
Attribute VB_Name = "ArrecEletronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
'Private qryGetArrecEletronica As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
'Private qryGetArrecEletronicaDuplicada As rdoQuery
'Private qryAtualizaDocumentoExcluido As rdoQuery
Private qryAtualizaArrecEletronica As rdoQuery
Private qryGetArrecEletronicaBarDuplicada As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Private bActivate As Boolean
Public Alterou As Boolean
Function DefineTipoDocto(ByVal sCodigoBarras As String) As Integer

    'Verificar se o documento é uma Arrecadação Eletrônica
    If (Mid(sCodigoBarras, 1, 1) = "8") Then
        'Determinar qual o tipo de Arrecadação
        Select Case (Mid(sCodigoBarras, 2, 1))
            Case "1", "5", "6"
                'TRIBUTOS MUNICIPAIS , ESTADUAIS , FEDERAIS e DPVAT-> Emitir Mensagem
                MsgBox "Este Documento é uma Arrecadação com Valor Indexado.", vbInformation, App.Title
                txtCodigo1.SetFocus
                DefineTipoDocto = 0
                Exit Function

            Case "2"
                'ÁGUA
                DefineTipoDocto = 20

            Case "3"
                'GÁS OU LUZ
                If Mid(sCodigoBarras, 17, 3) = "056" Or Mid(sCodigoBarras, 17, 3) = "057" Then
                    DefineTipoDocto = 21        'GÁS
                Else
                    DefineTipoDocto = 22        'LUZ
                End If

            Case "4"
                'TELEFONE
                DefineTipoDocto = 23

            Case Else
                'Sem Código
                MsgBox "Código de Barras inválido.", vbInformation, App.Title
                DefineTipoDocto = 0
        End Select
    Else
        MsgBox "Este Documento não é uma Arrecadação Eletrônica.", vbInformation, App.Title
        txtCodigo1.SetFocus
        DefineTipoDocto = 0
    End If
End Function
Sub ExibeValor()

  Dim svalor As Currency

  If Len(Trim(txtCodigo1.Text)) = 12 And Len(Trim(txtCodigo2.Text)) = 12 Then
    svalor = Mid(txtCodigo1.Text, 5, 7) & Mid(txtCodigo2.Text, 1, 4)
    txtValor.Text = svalor
  End If
End Sub
Function ValidaCodigoBarras(ByVal CodigoBarras As String) As Boolean

  On Error GoTo ValidaCodigoBarras_Err

  Dim sCodigoBarras As String
  Dim sSql As String
  Dim RsArrec As rdoResultset

  ValidaCodigoBarras = False

  'Verificar se o código de barras é válido
  If Not VerificaCodigo1 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    txtCodigo1.Text = ""
    
    txtCodigo1.SelStart = 0
    txtCodigo1.SelLength = Len(txtCodigo1.Text)

    txtCodigo1.SetFocus
    Exit Function
  End If

  If Not VerificaCodigo2 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    txtCodigo2.Text = ""
    
    txtCodigo2.SelStart = 0
    txtCodigo2.SelLength = Len(txtCodigo2.Text)
            
    txtCodigo2.SetFocus
    Exit Function
  End If

  If Not VerificaCodigo3 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    txtCodigo3.Text = ""
    
    txtCodigo3.SelStart = 0
    txtCodigo3.SelLength = Len(txtCodigo3.Text)
    
    txtCodigo3.SetFocus
    Exit Function
  End If

  If Not VerificaCodigo4 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    txtCodigo4.Text = ""
    
    txtCodigo4.SelStart = 0
    txtCodigo4.SelLength = Len(txtCodigo4.Text)
    txtCodigo4.SetFocus
    Exit Function
  End If

  'Verificar se este documento já foi gravado com valor diferente
  With qryGetArrecEletronicaBarDuplicada
    .rdoParameters(0).Direction = rdParamReturnValue
    .rdoParameters(1).Value = Geral.DataProcessamento
    .rdoParameters(2).Value = Mid(CodigoBarras, 1, 4)
    .rdoParameters(3).Value = Mid(CodigoBarras, 5, 11)
    .rdoParameters(4).Value = Mid(CodigoBarras, 16, 29)
    .rdoParameters(5).Value = Geral.Documento.IdDocto
    .Execute
  End With

  If qryGetArrecEletronicaBarDuplicada(0).Value = 1 Then
    'Encontrou outro documento com código de barras igual e valor diferente
    MsgBox "Já existe outro documento com o mesmo código de barras e valor diferente , por isso , este documento deve ser complementado como 'Arrecadação Convencional';", vbInformation + vbOKOnly, App.Title
    txtCodigo1.SetFocus
    Exit Function
  End If

  'Verificar se a Arrecadação é Convencional
  sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                  Mid(txtCodigo4.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)

  If VerificarArrecadacaoConvencional(sCodigoBarras) Then
    MsgBox "Este Documento é uma Arrecadação Convencional.", vbInformation, App.Title
    Exit Function
  End If

  ValidaCodigoBarras = True

  Exit Function

ValidaCodigoBarras_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Código de Barras.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdConfirmar_Click()

  If SalvaArrecEletronica Then
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

  bActivate = True

  Call AjustesIniciais

  Call PesquisaArrecEletronica

  Screen.MousePointer = vbDefault

  bActivate = False
End Sub
Private Function VerificaCodigo4() As Boolean

    Dim sArrec As String

    VerificaCodigo4 = False

    'Verifica se o campo foi preenchido por completo
    If Len(Trim(txtCodigo4.Text)) <> 12 Then Exit Function

    'Verifica se o campo está zerado
'    If Val(txtCodigo4.Text) = 0 Then Exit Function

    'Calculo do Modulo 10 para o campo 4
    If Not Modulo10(txtCodigo4.Text, 12) Then Exit Function

    'Só calcula este digito se não for o CDAE-RJ
    If ((Mid(txtCodigo1.Text, 1, 1) = "8") And (Mid(txtCodigo1.Text, 3, 1) = "6")) Then
        'Verifica se codigo de barras está batido atraves do 4º caracter
        sArrec = Mid(txtCodigo1.Text, 4, 1) + Mid(txtCodigo1.Text, 1, 3) + _
                 Mid(txtCodigo1.Text, 5, 7) + Mid(txtCodigo2.Text, 1, 11) + _
                 Mid(txtCodigo3.Text, 1, 11) + Mid(txtCodigo4.Text, 1, 11)

        If Not Modulo10Arrecadacao(sArrec, 44) Then Exit Function

    End If

    VerificaCodigo4 = True

End Function
Private Function VerificaCodigo3() As Boolean

    VerificaCodigo3 = False

    If Len(Trim(txtCodigo3.Text)) <> 12 Then Exit Function

    'Verifica se o campo está zerado
'    If Val(txtCodigo3.Text) = 0 Then Exit Function

    'Calculo do Modulo 10 para o campo 3
    If Not Modulo10(txtCodigo3.Text, 12) Then Exit Function

    VerificaCodigo3 = True
End Function
Private Function VerificaCodigo2() As Boolean

  VerificaCodigo2 = False

  If Len(Trim(txtCodigo2.Text)) <> 12 Then Exit Function

  'Verifica se o campo está zerado
  If Val(txtCodigo2.Text) = 0 Then Exit Function

  'Calculo do Modulo 10 para o campo 2
  If Not Modulo10(txtCodigo2.Text, 12) Then Exit Function

  VerificaCodigo2 = True
End Function
Private Function VerificaCodigo1() As Boolean

  VerificaCodigo1 = False

  If Len(Trim(txtCodigo1.Text)) <> 12 Then Exit Function

  'Verifica se o campo está zerado
  If Val(txtCodigo1.Text) = 0 Then Exit Function

  'Verifica se é mesmo uma concessionaria expressa em reais ou se CDAE-RJ
  If ((Mid(txtCodigo1.Text, 1, 1) = "8") And (Mid(txtCodigo1.Text, 3, 1) = "6")) Or _
     ((Mid(txtCodigo1.Text, 1, 1) = "8") And (Mid(txtCodigo1.Text, 2, 1) = "2") And (Mid(txtCodigo1.Text, 3, 1) = "4")) Then

     If Not Modulo10(txtCodigo1.Text, 12) Then Exit Function

  Else
    Exit Function
  End If

  VerificaCodigo1 = True
End Function
Function SalvaArrecEletronica() As Boolean

  On Error GoTo ERRO_SALVAARREC

  Dim sCodigoBarras As String
  Dim TipoDocto As Integer
  Dim strEncripta   As String

  SalvaArrecEletronica = False

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                    Mid(txtCodigo3.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)
  
    If Not ValidaCodigoBarras(sCodigoBarras) Then Exit Function

    'Verificar qual o tipo de documento
    TipoDocto = DefineTipoDocto(sCodigoBarras)

    If TipoDocto = 0 Then Exit Function

    'Verificar se o Documento pertence à outro Tipo
    If Val(Geral.Documento.TipoDocto) <> 20 And Val(Geral.Documento.TipoDocto) <> 21 And _
       Val(Geral.Documento.TipoDocto) <> 22 And Val(Geral.Documento.TipoDocto) <> 23 And _
       Val(Geral.Documento.TipoDocto) <> 0 Then
'       Val(Geral.Documento.TipoDocto) <> 27 And Val(Geral.Documento.TipoDocto) <> 0 Then

      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(TipoDocto, sCodigoBarras)
    If strEncripta = "" Then GoTo ERRO_SALVAARREC

    'Atualizar / Inserir Arrecadação Eletronica
    With qryAtualizaArrecEletronica
      .rdoParameters(0).Direction = rdParamReturnValue            'Parametro de Retorno
      .rdoParameters(1) = Geral.DataProcessamento                 'Data Proc.
      .rdoParameters(2) = Geral.Documento.IdDocto                 'IdDocto
      .rdoParameters(3) = sCodigoBarras                           'Codigo de Barras
      .rdoParameters(4) = Val(txtValor.Text) / 100                'Valor
      .rdoParameters(5) = TipoDocto                               'TipoDocto
      .rdoParameters(6) = strEncripta                             'Autenticacao digital
      .Execute
    End With

    If qryAtualizaArrecEletronica(0).Value = 2 Then
      'Documento Duplicado
      Geral.Documento.Status = "D"
    End If

    SalvaArrecEletronica = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(txtValor.Text) / 100
    Geral.Documento.Leitura = sCodigoBarras
    Geral.Documento.TipoDocto = TipoDocto
  End If

  Exit Function

ERRO_SALVAARREC:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Arrecadação Eletronica.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PesquisaArrecEletronica()

  On Error GoTo ERRO_PESQUISAARRECELETRONICA

  Dim sSql As String
  Dim RsArrec As rdoResultset

  txtCodigo1.Text = Mid(Geral.Documento.Leitura, 1, 11)
  txtCodigo1.Text = txtCodigo1.Text & DV10(txtCodigo1.Text)
  txtCodigo2.Text = Mid(Geral.Documento.Leitura, 12, 11)
  txtCodigo2.Text = txtCodigo2.Text & DV10(txtCodigo2.Text)
  txtCodigo3.Text = Mid(Geral.Documento.Leitura, 23, 11)
  txtCodigo3.Text = txtCodigo3.Text & DV10(txtCodigo3.Text)
  txtCodigo4.Text = Mid(Geral.Documento.Leitura, 34, 11)
  txtCodigo4.Text = txtCodigo4.Text & DV10(txtCodigo4.Text)
  
  If Geral.Documento.Status = "0" Then
    If Mid(txtCodigo2.Text, 9, 3) = "000" Then
        txtCodigo2.Text = String(12, "0")
    End If
      
    If Mid(txtCodigo3.Text, 9, 3) = "000" Then
        txtCodigo3.Text = String(12, "0")
    End If
    
    If Mid(txtCodigo4.Text, 9, 3) = "000" Then
        txtCodigo4.Text = String(12, "0")
    End If
  End If

  If Len(Trim(Geral.Documento.Leitura)) = 44 Then
    'Verificar se o DV está correto
    If VerificaCodigo1 Then
      If VerificaCodigo2 Then
        If VerificaCodigo3 Then
          If VerificaCodigo4 Then
            txtCodigo4.SelStart = 0
            txtCodigo4.SelLength = Len(txtCodigo4.Text)
            txtCodigo4.SetFocus
            Exit Sub
          End If
        End If
      End If
    End If
  End If

  txtValor.Text = ""
  txtCodigo1.SetFocus

  'Preencher os campos da Arrecadação , caso encontre
  'sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto
  'Set qryGetArrecEletronica = Geral.Banco.CreateQuery("", "{call GetArrecEletronica (" & sSql & ")}")

  'Set RsArrec = qryGetArrecEletronica.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  'If Not RsArrec.EOF Then
  '  txtValor.Text = RsArrec!Valor * 100
  'End If

  Exit Sub

ERRO_PESQUISAARRECELETRONICA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados da Arrecadação Eletrônica.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CamposOK() As Boolean

  'Primeiro Campo do Código de Barras
  If Len(Trim(txtCodigo1.Text)) = 0 Then
    MsgBox "Informe o Primeiro Campo do Código de Barras.", vbInformation, App.Title
    CamposOK = False
    txtCodigo1.SetFocus
    Exit Function
  End If

  'Segundo Campo do Código de Barras
  If Len(Trim(txtCodigo2.Text)) = 0 Then
    MsgBox "Informe o Segundo Campo do Código de Barras.", vbInformation, App.Title
    CamposOK = False
    txtCodigo2.SetFocus
    Exit Function
  End If

  'Terceiro Campo do Código de Barras
  If Len(Trim(txtCodigo3.Text)) = 0 Then
    MsgBox "Informe o Terceiro Campo do Código de Barras.", vbInformation, App.Title
    CamposOK = False
    txtCodigo3.SetFocus
    Exit Function
  End If

  'Quarto Campo do Código de Barras
  If Len(Trim(txtCodigo4.Text)) = 0 Then
    MsgBox "Informe o Quarto Campo do Código de Barras.", vbInformation, App.Title
    CamposOK = False
    txtCodigo4.SetFocus
    Exit Function
  End If

  If Val(txtValor.Text) = 0 Then
    MsgBox "Valor do documento está zerado, favor pagar como Arrecadação Convencional.", vbInformation, App.Title
    CamposOK = False
    txtCodigo1.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Sub AjustesIniciais()

  'Setando as Variáveis do RDO
'  Set qryGetArrecEletronica = Geral.Banco.CreateQuery("", "{? = call GetArrecEletronica (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
'  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaArrecEletronica = Geral.Banco.CreateQuery("", "{? = call AtualizaArrecEletronica (?,?,?,?,?,?)}")
  Set qryGetArrecEletronicaBarDuplicada = Geral.Banco.CreateQuery("", "{? = call GetArrecEletronicaBarDuplicada (?,?,?,?,?)}")
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

'  Set qryGetArrecEletronica = Nothing
  Set qryRemoveTipoDocumento = Nothing
'  Set qryAtualizaDocumentoExcluido = Nothing
  Set qryAtualizaArrecEletronica = Nothing
  'Set qryGetArrecEletronicaDuplicada = Nothing
End Sub
Private Sub txtCodigo1_Change()

  If Len(Trim(txtCodigo1.Text)) = txtCodigo1.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCodigo1_GotFocus()

  txtCodigo1.SelStart = 0
  txtCodigo1.SelLength = txtCodigo1.MaxLength
End Sub

Private Sub txtCodigo1_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtCodigo1_LostFocus()

  Call ExibeValor
End Sub

Private Sub txtCodigo2_Change()

  If Len(Trim(txtCodigo2.Text)) = txtCodigo2.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCodigo2_GotFocus()

  txtCodigo2.SelStart = 0
  txtCodigo2.SelLength = txtCodigo2.MaxLength
End Sub

Private Sub txtCodigo2_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtCodigo2_LostFocus()

  Call ExibeValor
End Sub
Private Sub txtCodigo3_Change()

  If Len(Trim(txtCodigo3.Text)) = txtCodigo3.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCodigo3_GotFocus()

  txtCodigo3.SelStart = 0
  txtCodigo3.SelLength = txtCodigo3.MaxLength
End Sub

Private Sub txtCodigo3_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtCodigo4_Change()

  If Len(txtCodigo4.Text) = txtCodigo4.MaxLength And bActivate = False Then
    Call cmdConfirmar_Click
  End If
End Sub

Private Sub txtCodigo4_GotFocus()

  txtCodigo4.SelStart = 0
  txtCodigo4.SelLength = txtCodigo4.MaxLength
End Sub

Private Sub txtCodigo4_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtValor_GotFocus()

  txtValor.SelStart = 0
  txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
