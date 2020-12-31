VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form ArrecValorIndexado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação do Codigo de Barras com Valor Referência"
   ClientHeight    =   1944
   ClientLeft      =   1284
   ClientTop       =   1308
   ClientWidth     =   9960
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1944
   ScaleWidth      =   9960
   Begin CURRENCYEDITLib.CurrencyEdit txtValor 
      Height          =   372
      Left            =   7680
      TabIndex        =   4
      Top             =   1476
      Width           =   1980
      _Version        =   65537
      _ExtentX        =   3492
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
      Left            =   2292
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
      Left            =   3984
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
      Left            =   600
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
      Left            =   5664
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
      Left            =   8892
      Picture         =   "ArrecValorIndexado.frx":0000
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
      Left            =   7248
      Picture         =   "ArrecValorIndexado.frx":030A
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
      Left            =   6420
      Picture         =   "ArrecValorIndexado.frx":0614
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
      Left            =   5592
      Picture         =   "ArrecValorIndexado.frx":091E
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
      Left            =   4764
      Picture         =   "ArrecValorIndexado.frx":0C28
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
      Left            =   3936
      Picture         =   "ArrecValorIndexado.frx":0F32
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
      Left            =   8064
      Picture         =   "ArrecValorIndexado.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   72
      Width           =   816
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Código de Barras com Valor Referência"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   768
      TabIndex        =   14
      Top             =   324
      Width           =   2892
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   204
      Picture         =   "ArrecValorIndexado.frx":1546
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
      Left            =   7776
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
      Left            =   600
      TabIndex        =   12
      Top             =   1140
      Width           =   3396
   End
End
Attribute VB_Name = "ArrecValorIndexado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryInserirConcessionariaSValor As rdoQuery      'Insere/Atualiza tabela CBIndex
Private qryAtualizaDocumentoExcluido As rdoQuery        'Atualiza Status = "D", Duplicidade = 1, Ocorrencia = 998
Private qryRemoveTipoDocumento As rdoQuery
Private qryGetArrecValorIndexadoBarDuplicada As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Function DefineTipoDocto(ByVal sCodigoBarras As String) As Integer

    'Verificar se o documento é uma Arrecadação Eletrônica
    If (Mid(sCodigoBarras, 1, 1) = "8") Then
        'Determinar qual o tipo de Arrecadação
        Select Case (Mid(sCodigoBarras, 2, 1))
            Case "1"
                'TRIBUTOS MUNICIPAIS
                DefineTipoDocto = 24

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

            Case "5"
                'TRIBUTOS
                If Val(Mid(sCodigoBarras, 17, 3)) <= 27 Then
                    'TRIBUTOS ESTADUAIS
                    DefineTipoDocto = 25
                Else
                    'TRIBUTOS FEDERAIS
                    DefineTipoDocto = 26
                End If

            Case "6"
                'DPVAT
                DefineTipoDocto = 25

            Case Else
                'Sem Código
                DefineTipoDocto = 0
        End Select
    Else
        DefineTipoDocto = 0
    End If
End Function
Function ValidaCodigoBarras() As Boolean

    Dim sCodigoBarras As String
    Dim sArrec        As String

    ValidaCodigoBarras = False

    'Verificar se o código de barras é válido
    If Not VerificaCodigo1 Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo1.Text = ""
        txtCodigo1.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo2 Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo2.Text = ""
        txtCodigo2.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo3 Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo3.Text = ""
        txtCodigo3.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo4 Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo4.Text = ""
        txtCodigo4.SetFocus
        Exit Function
    End If

    sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                    Mid(txtCodigo3.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)

    'Verificar se este documento já foi gravado com valor diferente
    With qryGetArrecValorIndexadoBarDuplicada
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = Mid(sCodigoBarras, 1, 4)
        .rdoParameters(3).Value = Mid(sCodigoBarras, 5, 11)
        .rdoParameters(4).Value = Mid(sCodigoBarras, 16, 29)
        .rdoParameters(5).Value = Geral.Documento.IdDocto
        .Execute
    End With

    If qryGetArrecValorIndexadoBarDuplicada(0).Value = 1 Then
        'Encontrou outro documento com código de barras igual e valor diferente
        MsgBox "Já existe outro documento com o mesmo código de barras e valor diferente , por isso , este documento deve ser complementado como 'Arrecadação Convencional';", vbInformation + vbOKOnly, App.Title
        txtCodigo1.SetFocus
        Exit Function
    End If

    'Verifica se codigo de barras está batido atraves do 4º caracter
    sArrec = Mid(txtCodigo1.Text, 4, 1) + Mid(txtCodigo1.Text, 1, 3) + _
             Mid(txtCodigo1.Text, 5, 7) + Mid(txtCodigo2.Text, 1, 11) + _
             Mid(txtCodigo3.Text, 1, 11) + Mid(txtCodigo4.Text, 1, 11)
           

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se Tributo Municipal - Mobiliario, colocar o valor'
    'contido no codigo de barras no campo valor e      '
    'desabilitar a digitação no campo valor            '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TributoComValorFixo Then
       
        ''''''''''''''''''''''''''''''''''
        'Pega o valor do codigo de barras'
        'xxx
        ''''''''''''''''''''''''''''''''''
        txtValor.Text = Mid(Left(txtCodigo1.Text, Len(txtCodigo1.Text) - 1) & Left(txtCodigo2.Text, Len(txtCodigo2.Text) - 1), 5, 11)
        txtValor.Locked = True
        
        'Verifica se documento de arrecadação está vencido
        If Mid(txtCodigo1.Text, 1, 3) = "816" And Mid(txtCodigo2.Text, 9, 2) = "27" Then
            If (Mid(txtCodigo2.Text, 11, 1) & Mid(txtCodigo3.Text, 1, 7)) < Geral.DataProcessamento Then
                MsgBox "DOCUMENTO VENCIDO.", vbInformation, App.Title
                txtCodigo1.SelStart = 0
                txtCodigo1.SelLength = Len(txtCodigo1.Text)
                txtCodigo1.SetFocus
                Exit Function
            End If
        End If
    Else
        txtValor.Locked = False
    End If

    If Not Modulo10Arrecadacao(sArrec, 44) Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo1.SelStart = 0
        txtCodigo1.SelLength = Len(txtCodigo1.Text)
        txtCodigo1.SetFocus
        Exit Function
    End If
    
    sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                    Mid(txtCodigo3.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)
    
    '* Verificar se docto é FGTS com código de barras
    If Mid(txtCodigo2.Text, 6, 3) = 107 Or Mid(txtCodigo2.Text, 6, 3) = 108 Or _
       Mid(txtCodigo2.Text, 6, 3) = 111 Or Mid(txtCodigo2.Text, 6, 3) = 112 Then
        MsgBox "Este Documento é um FGTS com código de barras.", vbInformation + vbOKOnly, App.Title
        Geral.Documento.Leitura = sCodigoBarras
        Exit Function
    End If
  
    'Verificar se a Arrecadação é Convencional
    If VerificarArrecadacaoConvencional(sCodigoBarras) Then
        MsgBox "Este Documento é uma Arrecadação Convencional.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End If

    ValidaCodigoBarras = True
End Function
Private Sub cmdConfirmar_Click()

    If Not SalvaDados Then Exit Sub
        
    Alterou = True
    Me.Hide

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
    Dim sArrec As String

    Call AjustesIniciais

    If Not VerificaCodigo1 Then
        txtCodigo1.SetFocus
        Exit Sub
    End If
    If Not VerificaCodigo2 Then
        txtCodigo2.SetFocus
        Exit Sub
    End If
    If Not VerificaCodigo3 Then
        txtCodigo3.SetFocus
        Exit Sub
    End If
    If Not VerificaCodigo4 Then
        txtCodigo4.SetFocus
        Exit Sub
    End If
    
    'Verifica se codigo de barras está batido atraves do 4º caracter
    sArrec = Mid(txtCodigo1.Text, 4, 1) + Mid(txtCodigo1.Text, 1, 3) + _
             Mid(txtCodigo1.Text, 5, 7) + Mid(txtCodigo2.Text, 1, 11) + _
             Mid(txtCodigo3.Text, 1, 11) + Mid(txtCodigo4.Text, 1, 11)

    If Not Modulo10Arrecadacao(sArrec, 44) Then
        txtCodigo1.SelStart = 0
        txtCodigo1.SelLength = Len(txtCodigo1.Text)
        txtCodigo1.SetFocus
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se Tributo Municipal - Mobiliario, colocar o valor'
        'contido no codigo de barras no campo valor e      '
        'desabilitar a digitação no campo valor            '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        If TributoComValorFixo Then
           
            ''''''''''''''''''''''''''''''''''
            'Pega o valor do codigo de barras'
            'xxx
            ''''''''''''''''''''''''''''''''''
            txtValor.Text = Mid(Left(txtCodigo1.Text, Len(txtCodigo1.Text) - 1) & Left(txtCodigo2.Text, Len(txtCodigo2.Text) - 1), 5, 11)
            txtValor.Locked = True
        Else
            txtValor.Locked = False
            txtValor.SetFocus
        End If
    End If
End Sub
Private Function VerificaCodigo4() As Boolean

    Dim sArrec As String

    VerificaCodigo4 = False

    If Len(Trim(txtCodigo4.Text)) <> 12 Then Exit Function
'    If Trim(txtCodigo4.Text) = String(12, "0") Then Exit Function

    'Calculo do Modulo 10 para o campo 4
    If Not Modulo10(txtCodigo4.Text, 12) Then Exit Function

    'Verifica se codigo de barras está batido atraves do 4º caracter
    sArrec = Mid(txtCodigo1.Text, 4, 1) + Mid(txtCodigo1.Text, 1, 3) + _
             Mid(txtCodigo1.Text, 5, 7) + Mid(txtCodigo2.Text, 1, 11) + _
             Mid(txtCodigo3.Text, 1, 11) + Mid(txtCodigo4.Text, 1, 11)

    If Not Modulo10Arrecadacao(sArrec, 44) Then Exit Function

    VerificaCodigo4 = True

End Function
Private Function VerificaCodigo3() As Boolean

  VerificaCodigo3 = False

  If Len(Trim(txtCodigo3.Text)) <> 12 Then Exit Function
'  If Trim(txtCodigo3.Text) = String(12, "0") Then Exit Function

  'Calculo do Modulo 10 para o campo 3
  If Not Modulo10(txtCodigo3.Text, 12) Then Exit Function

  VerificaCodigo3 = True
  
End Function
Private Function VerificaCodigo2() As Boolean

  VerificaCodigo2 = False

  If Len(Trim(txtCodigo2.Text)) <> 12 Then Exit Function
  If Trim(txtCodigo2.Text) = String(12, "0") Then Exit Function

  'Calculo do Modulo 10 para o campo 2
  If Not Modulo10(txtCodigo2.Text, 12) Then Exit Function

  VerificaCodigo2 = True
  
End Function
Private Function VerificaCodigo1() As Boolean

    VerificaCodigo1 = False

    If Len(Trim(txtCodigo1.Text)) <> 12 Then Exit Function
    If Trim(txtCodigo1.Text) = String(12, "0") Then Exit Function

    'Verifica se é uma concessionaria
    If (Mid(txtCodigo1, 1, 1) <> "8") Then Exit Function
  
    'Verifica se é mesmo uma concessionaria NÃO expressa em reais
    If (Mid(txtCodigo1, 1, 1) = "8") And (Mid(txtCodigo1, 3, 1) <> "6" And Mid(txtCodigo1, 3, 1) <> "7") Then Exit Function
    
    If Not Modulo10(txtCodigo1.Text, 12) Then Exit Function
        
    If txtValor.Locked And Not TributoComValorFixo Then txtValor.Locked = False

    VerificaCodigo1 = True
  
End Function
Function SalvaDados() As Boolean

On Error GoTo Err_SalvaDados
    
    Dim sCodigoBarras As String
    Dim TipoDocto As Integer
    Dim dValor As Double
    Dim bDuplicidade As Boolean
    Dim RetAgencia  As Integer
    Dim sDataVencto As String
    Dim strEncripta   As String
    
    SalvaDados = False

    'Verificar se todos os campos estão preenchidos
    If CamposOK Then
        If Not ValidaCodigoBarras Then Exit Function

        sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                        Mid(txtCodigo3.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)

        'Verificar qual o tipo de documento
        TipoDocto = DefineTipoDocto(sCodigoBarras)
        
        If Not (TipoDocto = 26 Or TipoDocto = 24 Or TipoDocto = 25) Then
            Beep
            MsgBox "Documento não identificado como Valor de Referência", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If

        'Verifica se Docto é ITBI ou Tributo 20 (Arrec. Dívida Ativa) ou Lixo
        If ((TipoDocto = 24 And Geral.DataProcessamento > 20030501) And (Mid(sCodigoBarras, 20, 2) = "56" Or Mid(sCodigoBarras, 20, 2) = "57")) Or _
            TipoDocto = 24 And (Mid(sCodigoBarras, 20, 2) = "31" Or _
            (Mid(sCodigoBarras, 20, 2) = "20" And InStr("0000*5889", Mid(sCodigoBarras, 16, 4)) > 0)) Then
            
            If Mid(sCodigoBarras, 20, 2) = "20" Then
                sDataVencto = CalculaVenctoDividaAtiva(sCodigoBarras)
            ElseIf Mid(sCodigoBarras, 20, 2) = "56" Or Mid(sCodigoBarras, 20, 2) = "57" Then
                sDataVencto = Format(Mid(sCodigoBarras, 34, 2) + "/" + Mid(sCodigoBarras, 36, 2) + "/" + Mid(sCodigoBarras, 38, 2), "ddmmyyyy")
            Else
                'Verifica se Docto está dentro da data de Vencimento
                sDataVencto = Mid(sCodigoBarras, 28, 2) + Mid(sCodigoBarras, 26, 2) + Mid(sCodigoBarras, 22, 4)
            End If
            
            'Verifica se Data Válida
            If Not DataOk(sDataVencto) Then
                Beep
                MsgBox "Erro no Código de Barras. Verifique.", vbInformation + vbOKOnly, App.Title
                Exit Function
            End If
            
            'Validar Agencia
            If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, sDataVencto, True, True) Then
                Exit Function
            End If
            
        End If
        
        'Inicia Transação
        Geral.Banco.BeginTrans

        'Verificar se o Documento pertence à outro Tipo
        If Val(Geral.Documento.TipoDocto) <> 8 And Val(Geral.Documento.TipoDocto) <> 9 And _
           Val(Geral.Documento.TipoDocto) <> 24 And Val(Geral.Documento.TipoDocto) <> 25 And _
           Val(Geral.Documento.TipoDocto) <> 26 And Val(Geral.Documento.TipoDocto) <> 0 Then
    
          With qryRemoveTipoDocumento
            .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
            .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
            .Execute
          End With
        End If
        
        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(TipoDocto, sCodigoBarras)
        If strEncripta = "" Then GoTo Exit_SalvaDados

        dValor = CDbl(txtValor.Text) / 100
        'Salva dados na Tabela CBIndex
        With qryInserirConcessionariaSValor
            .rdoParameters(1) = Geral.DataProcessamento         'Data Processamento
            .rdoParameters(2) = Geral.Documento.IdDocto         'IdDocto
            .rdoParameters(3) = sCodigoBarras                   'Código de Barras
            .rdoParameters(4) = dValor                          'Valor
            .rdoParameters(5) = strEncripta                     'Autenticacao digital
            .Execute
    
            'Verifica se ocorreu erro na atualização
            If .rdoParameters(0).Value <> 0 Then GoTo Exit_SalvaDados
        
        End With
    
        'Atualiza tabela Documento
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , TipoDocto, sCodigoBarras, , , , dValor) Then
            GoTo Exit_SalvaDados
        End If

        'Verifica se houve ocorrencia de duplicidade para o campo Leitura
        If bDuplicidade Then
            Geral.Documento.Duplicidade = 1
            If Not AtualizaDocumentoExcluido(Geral.Documento.IdDocto) Then
                GoTo Exit_SalvaDados
            End If
        Else
            Geral.Documento.Duplicidade = 0
        End If
            
        If Geral.Documento.Duplicidade = 0 Then
            Geral.Documento.Status = "1"
        Else
            Geral.Documento.Status = "D"
        End If

        'Atualiza status do documento
        bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
        If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , , , Geral.Documento.Status) Then
            GoTo Exit_SalvaDados
        End If
        
        'Finaliza rotina
        SalvaDados = True
    
        'Atualiza variaveis globais de Documento
        Geral.Documento.TipoDocto = TipoDocto
        Geral.Documento.Leitura = sCodigoBarras
        Geral.Documento.ValorTotal = dValor
        
        'Finaliza Transação
        Geral.Banco.CommitTrans
        
    End If

    Exit Function

Exit_SalvaDados:
    Geral.Banco.RollbackTrans
    MsgBox "Não foi possível Incluir/atualiza Código de Barras com Valor Indexado.", vbCritical + vbOKOnly, App.Title
    CmdSair_Click
    Exit Function

Err_SalvaDados:
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Erro ao Atualizar Dados do Documento.", Err, rdoErrors)
        Case vbCancel
            Alterou = False
            Me.Hide
        Case vbRetry
    End Select

End Function
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

  'Quarto Campo do Código de Barras
  If Val(txtValor.Text) = 0 Then
    MsgBox "Informe o Valor do documento.", vbInformation, App.Title
    CamposOK = False
    txtValor.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Sub AjustesIniciais()

    'Insere/Altera campos da tabela CBIndex
    Set qryInserirConcessionariaSValor = Geral.Banco.CreateQuery("", "{? = call InserirConcessionariaSValor (?,?,?,?,?)}")
    Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoExcluido (?,?,?,?,?)}")
    Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
    Set qryGetArrecValorIndexadoBarDuplicada = Geral.Banco.CreateQuery("", "{? = call GetArrecValorIndexadoBarDuplicada (?,?,?,?,?)}")

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

Private Sub Form_Load()
    
    If Geral.Documento.ValorTotal <> 0 Then
        txtValor.Text = Geral.Documento.ValorTotal
    Else
        txtValor.Text = ""
    End If
    
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
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  qryAtualizaDocumentoExcluido.Close
  qryInserirConcessionariaSValor.Close
  qryRemoveTipoDocumento.Close
  
  
  Set qryInserirConcessionariaSValor = Nothing
  Set qryAtualizaDocumentoExcluido = Nothing
  Set qryRemoveTipoDocumento = Nothing
  
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
  
  If Len(Trim(txtCodigo4.Text)) = txtCodigo4.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If

End Sub

Private Sub txtCodigo4_GotFocus()

  txtCodigo4.SelStart = 0
  txtCodigo4.SelLength = txtCodigo4.MaxLength
End Sub

Private Sub txtCodigo4_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub txtCodigo4_Validate(Cancel As Boolean)


'    If Not VerificaCodigo4 Then
'        MsgBox "Código de Barras inválido.", vbExclamation
'        Cancel = True
'        Exit Sub
'    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se Tributo Municipal - Mobiliario, colocar o valor'
    'contido no codigo de barras no campo valor e      '
    'desabilitar a digitação no campo valor            '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TributoComValorFixo Then
       
        ''''''''''''''''''''''''''''''''''
        'Pega o valor do codigo de barras'
        'xxx
        ''''''''''''''''''''''''''''''''''
        txtValor.Text = Mid(Left(txtCodigo1.Text, Len(txtCodigo1.Text) - 1) & Left(txtCodigo2.Text, Len(txtCodigo2.Text) - 1), 5, 11)
        txtValor.Locked = True
    Else
        txtValor.Locked = False
    End If

End Sub

Private Sub txtValor_GotFocus()
    
    With Me.txtValor
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
  
End Sub

Private Function AtualizaDocumentoExcluido(ByVal IdDocto As Long) As Boolean
    On Error GoTo ErroExclusao
    rdoErrors.Clear
    
    AtualizaDocumentoExcluido = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaDocumentoExcluido
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = "D" ' status
        .rdoParameters(4) = 1   ' duplicidade
        .rdoParameters(5) = 998 ' ocorrencia
        .Execute
        If .rdoParameters(0) <> 0 Then
            GoTo ErroExclusao
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroExclusao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Function

Private Function TributoComValorFixo() As Boolean
    
    If ((Mid(txtCodigo1.Text, 1, 3) = "817") And ( _
       (Mid(txtCodigo2.Text, 9, 2) = "43") Or _
       (Mid(txtCodigo2.Text, 9, 2) = "52") Or _
       (Mid(txtCodigo2.Text, 9, 2) = "23"))) _
       Or _
       ((Mid(txtCodigo1.Text, 1, 3) = "816") And _
       (Mid(txtCodigo2.Text, 9, 2) = "23") Or _
       (Mid(txtCodigo2.Text, 9, 2) = "27")) _
       Or _
       (Left(txtCodigo1.Text, 1) = "8" And Mid(txtCodigo1.Text, 3, 1) = "6") _
       Then
       
        TributoComValorFixo = True
    Else
        TributoComValorFixo = False
    End If

End Function
Private Function CalculaVenctoDividaAtiva(ByVal sCodigoBarras As String) As String

Dim iDiasCorrentes As Integer, iDigAnoVecto As Integer, dDataVecto As Date
Dim iAnoMovto As Integer, iDigAnoMovto As Integer

    'Verifica se cálculo da data de vencto é no formato Juliano
    If Mid(sCodigoBarras, 16, 4) = "5889" Then
        'Obtem o Nr de Dias correntes
        iDiasCorrentes = Val(Mid(sCodigoBarras, 23, 3))
        'Obtem o último dígito do ano de vecto
        iDigAnoVecto = Mid(sCodigoBarras, 22, 1)
        
        'Obtem Ano do movimento
        iAnoMovto = Val(Left(Geral.DataProcessamento, 4))
        
        'Soma-se o ano de movimento até igualar com o ano de vecto do tributo para
        'obter todos dígitos do ano
        iDigAnoMovto = Val(Right(CStr(iAnoMovto), 1))
        
        While iDigAnoMovto <> iDigAnoVecto
            If iDigAnoVecto = 0 Then
                If iDigAnoMovto > 7 Then
                    iAnoMovto = iAnoMovto + 1
                Else
                    iAnoMovto = iAnoMovto - 1
                End If
            Else
                If iDigAnoMovto > 7 Then
                    iAnoMovto = iAnoMovto + 1
                Else
                    iAnoMovto = iAnoMovto - 1
                End If
            End If
            iDigAnoMovto = Val(Right(CStr(iAnoMovto), 1))
        Wend
        
        'Obtem o Dia e Mês de vencto conforme dias correntes
        dDataVecto = CVDate("01/01/" + CStr(iAnoMovto)) + (iDiasCorrentes - 1)
        
        CalculaVenctoDividaAtiva = Right("00" + CStr(Day(dDataVecto)), 2) & _
                                    Right("00" + CStr(Month(dDataVecto)), 2) & _
                                    CStr(Year(dDataVecto))
    Else
        'Obtem a data de Vencimento no formato padrão
        CalculaVenctoDividaAtiva = Mid(sCodigoBarras, 26, 2) + Mid(sCodigoBarras, 24, 2) + Left(Geral.DataProcessamento, 2) + Mid(sCodigoBarras, 22, 2)
    End If

End Function
