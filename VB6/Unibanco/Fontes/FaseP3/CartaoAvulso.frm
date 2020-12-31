VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form CartaoAvulso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Cartão Avulso"
   ClientHeight    =   2856
   ClientLeft      =   -480
   ClientTop       =   828
   ClientWidth     =   10476
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2856
   ScaleWidth      =   10476
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   800
      Left            =   8640
      Picture         =   "CartaoAvulso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   800
      Left            =   9504
      Picture         =   "CartaoAvulso.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   800
      Left            =   5184
      Picture         =   "CartaoAvulso.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   96
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   800
      Left            =   4320
      Picture         =   "CartaoAvulso.frx":091E
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
      Left            =   6048
      Picture         =   "CartaoAvulso.frx":0C28
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
      Left            =   6912
      Picture         =   "CartaoAvulso.frx":0F32
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
      Left            =   7776
      Picture         =   "CartaoAvulso.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   96
      Width           =   852
   End
   Begin VB.Frame fraCartao 
      Height          =   1788
      Left            =   96
      TabIndex        =   15
      Top             =   960
      Width           =   10236
      Begin CURRENCYEDITLib.CurrencyEdit txtDespReais 
         Height          =   348
         Left            =   4224
         TabIndex        =   16
         Top             =   480
         Width           =   1980
         _Version        =   65537
         _ExtentX        =   3492
         _ExtentY        =   614
         _StockProps     =   93
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
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Left            =   7152
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1980
      End
      Begin VB.TextBox txtCartao 
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
         Left            =   912
         MaxLength       =   16
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   2280
      End
      Begin CURRENCYEDITLib.CurrencyEdit txtDespDolar 
         Height          =   348
         Left            =   4224
         TabIndex        =   17
         Top             =   1200
         Width           =   1980
         _Version        =   65537
         _ExtentX        =   3492
         _ExtentY        =   614
         _StockProps     =   93
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
      Begin CURRENCYEDITLib.CurrencyEdit txtAntSaque 
         Height          =   348
         Left            =   7152
         TabIndex        =   18
         Top             =   480
         Width           =   1980
         _Version        =   65537
         _ExtentX        =   3492
         _ExtentY        =   614
         _StockProps     =   93
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
      Begin VB.Label lblDespSaque 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rec. antecipado de Saque"
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
         Left            =   7140
         TabIndex        =   4
         Top             =   240
         Width           =   2220
      End
      Begin VB.Label lblDespReais 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas em Reais"
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
         Left            =   4224
         TabIndex        =   2
         Top             =   240
         Width           =   1668
      End
      Begin VB.Label lblDespDolar 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas em Dólar"
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
         Left            =   4224
         TabIndex        =   3
         Top             =   960
         Width           =   1644
      End
      Begin VB.Label lblCartao 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Cartão"
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
         Left            =   912
         TabIndex        =   0
         Top             =   240
         Width           =   1548
      End
      Begin VB.Label lblDespTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total"
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
         Left            =   7140
         TabIndex        =   5
         Top             =   960
         Width           =   912
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de Cartão de Crédito Avulso"
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
      Left            =   864
      TabIndex        =   14
      Top             =   360
      Width           =   3204
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   240
      Picture         =   "CartaoAvulso.frx":1546
      Top             =   240
      Width           =   384
   End
End
Attribute VB_Name = "CartaoAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variavel de retorno informando se Cancelou ou Alterou
Public Alterou As Boolean

Dim bAlterou As Boolean
Dim sPosicaoErro As String
Dim sMascValor As String
Private mForm As Form
Public AlteraValor As Boolean

Private Type tpModulo
    qryInserirCartao As rdoQuery
    qryGetCartaoAvulso As rdoQuery
    qryVerificaBinCartao As rdoQuery
    qryRemoveTipoDocumento As rdoQuery
End Type

Private Modulo As tpModulo

Private Function VerificarCartao() As Boolean

    Dim sCodigoBandeira As String

    If Len(txtCartao.Text) = 0 Or Val(txtCartao.Text) = 0 Then
        MsgBox "O preenchimento do número do cartão é obrigatório !", vbExclamation + vbOKOnly, App.Title
        GoTo Sair
    Else
        
'        'Se Número do Cartão é diferente de 14 dígitos (Dinners), formata com 16 dígitos (Credicard, Visa,..)
'        If Len(Trim(txtCartao.Text)) <> 14 Then
'            If Len(Trim(txtCartao.Text)) <> 16 Then txtCartao.Text = Format(txtCartao.Text, "0000000000000000")
'        End If

        'Verifica Código BIN do Cartão de Crédito
        If Not ConsisteBin(txtCartao.Text, sCodigoBandeira) Then
            MsgBox "Código do Cartão não aceito pelo Unibanco. Verifique!", vbExclamation + vbOKOnly, App.Title
            GoTo Sair
        End If
        
        If sCodigoBandeira = "00" Or sCodigoBandeira = "01" Then
            'Verifica Número do cartão
            If Not Modulo10(txtCartao.Text, 16) Then
                MsgBox "Número do cartão não confere. Verifique!", vbExclamation + vbOKOnly, App.Title
                GoTo Sair
            End If
        
        ElseIf sCodigoBandeira = "02" Then  'DINNERS
            'Verifica Número do cartão
            If Not Modulo10Dinners(txtCartao.Text) Then
                MsgBox "Número do cartão não confere. Verifique!", vbExclamation + vbOKOnly, App.Title
                GoTo Sair
            End If
        
        ElseIf sCodigoBandeira = "03" Then  'CREDICARD
            'Verifica Número do cartão
            If Not Modulo11Credicard(txtCartao.Text) Then
                MsgBox "Número do cartão não confere. Verifique!", vbExclamation + vbOKOnly, App.Title
                GoTo Sair
            End If
        Else
            MsgBox "Cartão não aceito pelo Unibanco. Verifique!", vbExclamation + vbOKOnly, App.Title
            GoTo Sair
        End If
        
        VerificarCartao = True
    
    End If
    
    Exit Function
    
Sair:
    VerificarCartao = False
    txtCartao_GotFocus
    txtCartao.SetFocus

End Function


Private Function VerificarTudo() As Boolean
    
    VerificarTudo = False
        
    'Valida número do cartão
    If VerificarCartao = True Then
        
        'Valida o valor despesas em reais
        If Len(Trim(txtDespReais.Text)) = 0 Or Val(Desformata_Valor(txtDespReais.Text)) = 0 Then
            MsgBox "O Valor das Despesas em Reais deve ser informado!", vbExclamation + vbOKOnly, App.Title
            VerificarTudo = False
            txtDespReais.SetFocus
            Exit Function
        Else
            
            'Valida o campo valor
            If Val(Desformata_Valor(txtDespReais.Text)) + Val(Desformata_Valor(txtDespDolar.Text)) + Val(Desformata_Valor(txtAntSaque.Text)) = 0 Then
                MsgBox "Digite o Valor do Documento.", vbExclamation + vbOKOnly, App.Title
                VerificarTudo = False
                Exit Function
            Else
                VerificarTudo = True
            End If
        End If
    End If

End Function

Private Sub cmdConfirmar_Click()
    
    txtAntSaque_KeyPress (vbKeyReturn)
    
End Sub

Private Sub cmdFrenteVerso_Click()
    
    mForm.cmdFrenteVerso_Click
    bAlterou = True
    
End Sub

Private Sub cmdInverteCor_Click()
    
    mForm.cmdInverteCor_Click
    bAlterou = True
    
End Sub

Private Sub cmdRotacao_Click()
    
    mForm.cmdRotacao_Click
    bAlterou = True
    
End Sub

Private Sub CmdSair_Click()
    
    Alterou = False
    Me.Hide
    
End Sub

Private Sub cmdZoomMais_Click()
    
    mForm.cmdZoomMais_Click
    bAlterou = True
    
End Sub

Private Sub cmdZoomMenos_Click()
    
    mForm.cmdZoomMenos_Click
    bAlterou = True
    
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
'                MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'            Case 3 'Agência Fechada
'                MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'        End Select
'        If iErroData = 2 Or iErroData = 3 Then
'            CmdSair_Click
'            Exit Sub
'        End If
'    End If
    
    'Ler dados da tabela Cartão Avulso, pode ou não existir registro
    If Not LerDadosCartao() Then CmdSair_Click: Exit Sub
    
    'Se form chamador é diferente de Complementação, desabilita controle de entrada <> de Valores
    If AlteraValor Then
        txtCartao.TabStop = False
        txtCartao.Locked = True
        txtCartao.ForeColor = vbBlack
        txtCartao.BackColor = G_ColorGray
        txtDespReais.SetFocus
    Else
        txtCartao.SetFocus
    End If
    
    bAlterou = False
    
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
  End Select
End Sub

Private Sub Form_Load()

    txtCartao.Text = ""
    txtValor.Text = ""
    Alterou = False
    
    txtValor.ForeColor = vbBlack
    txtValor.BackColor = G_ColorGray
    txtValor.Locked = True
    
    sMascValor = "###,###,##0.00  "
    
    With Modulo
        'Cria a query para a gravação dos dados do cartão
        Set .qryInserirCartao = Geral.Banco.CreateQuery("", "{? = call InserirCartaoAvulso (?,?,?,?,?,?,?,?)}")
            'Parametros: (1)-Data Processamento (2)-IdDocto (3)-Numero Cartao (4)-Desp.Reais (5)-Desp.Dolar (6)-Ant.Saque  (7)-Valor
            .qryInserirCartao.rdoParameters(0).Direction = rdParamReturnValue
            
        'Cria a query para Ler dados do cartão
        Set .qryGetCartaoAvulso = Geral.Banco.CreateQuery("", "{? = call GetCartaoAvulso (?,?)}")
            'Parametros (1)-Data Processamento (2)-IdDocto
            .qryGetCartaoAvulso.rdoParameters(0).Direction = rdParamReturnValue
            
        'Valida Código BIN Cartão
        Set .qryVerificaBinCartao = Geral.Banco.CreateQuery("", "{? = call VerificaBinCartao (?)}")
            'Parametros (1)-Código BIN do Cartão
            .qryVerificaBinCartao.rdoParameters(0).Direction = rdParamReturnValue
        
        Set .qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
        
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    With Modulo
        .qryGetCartaoAvulso.Close
        .qryInserirCartao.Close
        .qryVerificaBinCartao.Close
        .qryRemoveTipoDocumento.Close
    End With
    
End Sub

Private Sub txtAntSaque_GotFocus()

    With txtAntSaque
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtAntSaque_KeyPress(KeyAscii As Integer)
    
Dim dValor As Double
Dim bDuplicidade As Boolean
Dim strEncripta   As String

On Error GoTo Err_SalvaDados

    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If KeyAscii = 13 Then
        
        txtValor.Text = Format((Val(txtDespReais.Text) + Val(txtDespDolar.Text) + Val(txtAntSaque.Text)) / 100, sMascValor)
        
        If VerificarTudo = True Then
            
            sPosicaoErro = "InsCartaoAvulso"
            dValor = CDbl(txtValor.Text)
            
            'Inicia Transação
            Geral.Banco.BeginTrans
    
            'Verificar se o Documento pertence à outro Tipo
            If Geral.Documento.TipoDocto <> 36 And Geral.Documento.TipoDocto <> 0 Then
                With Modulo.qryRemoveTipoDocumento
                  .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                  .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
                  .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
                  .Execute
                End With
            End If

            'Atualiza campo Autenticação Digital
            strEncripta = G_EncriptaBO(etpdocCartaoAvulso, txtCartao.Text)
            If strEncripta = "" Then GoTo Exit_SalvaDados

            With Modulo.qryInserirCartao
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = Geral.Documento.IdDocto
                .rdoParameters(3) = txtCartao.Text
                .rdoParameters(4) = (Val(txtDespReais.Text) / 100)         'Desp. Reais
                .rdoParameters(5) = (Val(txtDespDolar.Text) / 100)         'Desp. Dolar
                .rdoParameters(6) = (Val(txtAntSaque.Text) / 100)          'Vlr. Saque
                .rdoParameters(7) = dValor                                 'Valor Total
                .rdoParameters(8) = strEncripta                             'Autenticacao digital
                .Execute

                If .rdoParameters(0).Value <> 0 Then GoTo Exit_SalvaDados
            
            End With
            
            Geral.Documento.TipoDocto = etpdocCartaoAvulso
            Geral.Documento.ValorTotal = dValor ' / 100
            Geral.Documento.Leitura = ""        'Não há necessidade de guardar campo leitura
            
            'Atualiza tabela Documento
            bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
            If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , etpdocCartaoAvulso, , , , , dValor) Then
                GoTo Exit_SalvaDados
            End If
            
            'Finaliza Transação
            Geral.Banco.CommitTrans
            
            Alterou = True
            Me.Hide
        
        End If
    
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
    
    bAlterou = True
    Exit Sub

Exit_SalvaDados:
    Alterou = False
    Geral.Banco.RollbackTrans
    MsgBox "Não foi possível Incluir/atualiza dados do Cartão de Crédito.", vbCritical + vbOKOnly, App.Title
'    cmdSair_Click
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

Private Sub txtAntSaque_LostFocus()
    
   txtValor.Text = Format((Val(txtDespReais.Text) + Val(txtDespDolar.Text) + Val(txtAntSaque.Text)) / 100, sMascValor)

End Sub

Private Sub txtCartao_GotFocus()
    
    With txtCartao
        .SelStart = 0
        .SelLength = .MaxLength
    End With

End Sub
Private Sub txtCartao_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If (KeyAscii = 13) Then
        If VerificarCartao = True Then
            txtDespReais.SetFocus
        End If
        
''''''''''
'Leda desabilitou para validar através da função
'20/03/2000
'        If Len(txtCartao) = 0 Then
'            MsgBox "O preenchimento deste campo é obrigatório !", vbExclamation + vbOKOnly, "Cartão Crédito Avulso"
'            txtCartao = ""
'            txtCartao.SetFocus
'        Else
'            If Len(Trim(txtCartao)) < 16 Then
'                txtCartao = Format(txtCartao, "0000000000000000")
'            End If
'            ''''''''''''''''''''''''
'            ' Calculo do Modulo 10 '
'            ''''''''''''''''''''''''
'            If Not Modulo10(txtCartao, 16) Then
'                MsgBox "Digito não confere. Tente novamente !", vbExclamation + vbOKOnly, "Cartão Crédito Avulso"
'                txtCartao.SetFocus
'            Else
'                txtDespReais.SetFocus
'            End If
'        End If
'''''''''''

    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub

Private Sub txtDespDolar_GotFocus()

    With txtDespDolar
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtDespDolar_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If KeyAscii = 13 Then
        txtValor.Text = Format((Val(txtDespReais.Text) + Val(txtDespDolar.Text) + Val(txtAntSaque.Text)) / 100, sMascValor)
        txtAntSaque.SetFocus
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub


Private Sub txtDespDolar_LostFocus()
    
        txtValor.Text = Format((Val(txtDespReais.Text) + Val(txtDespDolar.Text) + Val(txtAntSaque.Text)) / 100, sMascValor)

End Sub

Private Sub txtDespReais_GotFocus()

    With txtDespReais
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtDespReais_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If KeyAscii = 13 Then
        If Len(Trim(txtDespReais.Text)) = 0 Or Val(Desformata_Valor(txtDespReais.Text)) = 0 Then
            MsgBox "O Valor das Despesas em Reais deve ser informado!", vbExclamation + vbOKOnly, Caption
            txtDespReais.SetFocus
        Else
            txtDespDolar.SetFocus
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub

Private Sub txtDespReais_LostFocus()
        
        txtValor.Text = Format((Val(txtDespReais.Text) + Val(txtDespDolar.Text) + Val(txtAntSaque.Text)) / 100, sMascValor)

End Sub

Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

Private Function LerDadosCartao() As Boolean

Dim rstModulo As rdoResultset

On Error GoTo Err_LerDadosCartao
    
    LerDadosCartao = False
    
    With Modulo.qryGetCartaoAvulso
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto
        Set rstModulo = .OpenResultset(rdOpenStatic)
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível ler dados de cartão de Crédito.", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If
        
        If Not rstModulo.EOF() Then
            'Atualiza Variáveis globais
            Geral.Documento.Leitura = Trim(rstModulo!Cartao)
            Geral.Documento.ValorTotal = (rstModulo!Valor * 100)
            'Atualiza dados de entrada
            txtCartao.Text = Trim(rstModulo!Cartao)
            txtDespReais.Text = (rstModulo!DespReais * 100)
            txtDespDolar.Text = (rstModulo!DespDolar * 100)
            txtAntSaque.Text = (rstModulo!AntSaque * 100)
            txtValor.Text = Format(rstModulo!Valor, sMascValor)
        Else
            'Atualiza Variáveis globais
            Geral.Documento.Leitura = ""
            Geral.Documento.ValorTotal = 0
            'Atualiza dados de entrada
            txtCartao.Text = ""
            txtDespReais.Text = 0
            txtDespDolar.Text = 0
            txtAntSaque.Text = 0
            txtValor.Text = Format(0, sMascValor)
        End If
    End With
    
    LerDadosCartao = True

Exit_LerDadosCartao:
    Set rstModulo = Nothing
    Exit Function
    
Err_LerDadosCartao:

    Select Case TratamentoErro("Não foi possível ler dados referentes ao Cartão de Crédito !", Err, rdoErrors, False)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    GoTo Exit_LerDadosCartao

End Function

Private Function ConsisteBin(ByVal sCdCartao As String, ByRef sCodigoBandeira As String) As Boolean
'Parâmetro: (sCdCartao) - Seis caracteres iniciais do Número do Cartão
'
'Retorno:   (0)- Sucesso
'           (1)- Erro no SQL
Dim rstBandeira As rdoResultset

On Error GoTo Err_ConsisteBin

    ConsisteBin = False
    
    With Modulo.qryVerificaBinCartao
        .rdoParameters(1) = Left(sCdCartao, 6)
        Set rstBandeira = .OpenResultset(rdOpenStatic)
        
        'Verifica se ocorreu erro no SQL
        If .rdoParameters(0).Value = 1 Then GoTo Err_ConsisteBin
        
        'Verifica se existe Código BIN do Cartão
        If rstBandeira.EOF Then Exit Function
        
        'Retorna código de Bandeira
        sCodigoBandeira = rstBandeira!crefsbandei
    End With
    
    ConsisteBin = True
    Exit Function

Err_ConsisteBin:
    Select Case TratamentoErro("Não foi possível Verificar Número do Cartão.( VerifBinCartao )", Err, rdoErrors)
        Case vbCancel
            Alterou = False
            Me.Hide
        Case vbRetry
    End Select

End Function
Public Function Modulo10Dinners(ByVal base_calculo As String) As Boolean
   
   '---------------------------------------------------------------------------------
   'Número do cartão:                       9 9 9 9  9 9 9 9  9 9 9 9  9 9 9 D
   '                                        x x x x  x x x x  x x x x  x x x
   'Peso:                                   2 1 2 1  2 1 2 1  2 1 2 1  2 1 2
   '---------------------------------------------------------------------------------
    
    Dim Resultado As Integer, peso As Integer, soma As Integer, resto As Integer, i  As Integer, digito As Integer
   
    Modulo10Dinners = False   'Digito do campo 1 não confere
    
    peso = 2
    Resultado = 0
   
    For i = (Len(base_calculo) - 1) To 1 Step -1
        soma = Val(Mid(base_calculo, i, 1)) * peso
   
        If soma > 9 Then soma = soma - 9
        
        Resultado = Resultado + soma
   
        If peso = 2 Then
            peso = 1
        Else
            peso = 2
        End If
   Next
   
    resto = Resultado Mod 10
    If resto = 0 Then
        digito = 0
    Else
        digito = (10 - Fix(resto))
    End If
    
   '*** Verifica se dígito confere ***
   If Val(Right(base_calculo, 1)) = digito Then Modulo10Dinners = True  'digito do campo 1 confere
   
End Function

Public Function Modulo11Credicard(ByVal base_calculo As String) As Boolean

   '---------------------------------------------------------------------------------
   'Número do cartão:                       9 9 9 9  9 9 9 9  9 9 9 9  9 9 D D
   '                                                 x x x x  x x x x  x x
   'Peso:                                            5 4 3 2  7 6 5 4  3 2
   'Dígito Calculado no Modulo 11:                                         D
   'Dígito Calculado no Modulo 10:                                           D
   '---------------------------------------------------------------------------------
    
    Dim Resultado As Integer, peso As Integer, soma As Integer, resto As Integer, i  As Integer, digito As Integer
   
    Modulo11Credicard = False   'Digito não confere
    
    peso = 5
    Resultado = 0
   
    For i = 5 To (Len(base_calculo) - 2)
        soma = Val(Mid(base_calculo, i, 1)) * peso
   
        Resultado = Resultado + soma
   
        peso = peso - 1
        If peso < 2 Then peso = 7
   Next
   
    resto = Resultado Mod 11
    If resto <> 0 Then
        digito = 11 - resto
    ElseIf resto = 10 Then
        digito = 0
    Else
        digito = resto
    End If
    
    '*** Verifica se dígito MODULO11 confere ***
    If Val(Mid(base_calculo, 15, 1)) <> digito Then Exit Function
   
    '*** Verifica se dígito MODULO10 confere ***
    If Not Modulo10Dinners(base_calculo) Then Exit Function
   
    Modulo11Credicard = True  'digito confere

End Function

