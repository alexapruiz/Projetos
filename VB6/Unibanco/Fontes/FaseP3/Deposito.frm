VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form Deposito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Depósitos"
   ClientHeight    =   3264
   ClientLeft      =   1320
   ClientTop       =   1356
   ClientWidth     =   8808
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3264
   ScaleWidth      =   8808
   Begin VB.Frame Frame1 
      Height          =   3240
      Left            =   36
      TabIndex        =   16
      Top             =   -36
      Width           =   8676
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   6924
         Picture         =   "Deposito.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   216
         Width           =   804
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificação"
         Height          =   984
         Left            =   132
         TabIndex        =   25
         Top             =   1020
         Width           =   5124
         Begin VB.TextBox txtCMC71 
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
            Left            =   168
            MaxLength       =   8
            TabIndex        =   0
            Top             =   480
            Width           =   1068
         End
         Begin VB.TextBox txtCMC73 
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
            Left            =   2976
            MaxLength       =   12
            TabIndex        =   2
            Top             =   480
            Width           =   1584
         End
         Begin VB.TextBox txtCMC72 
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
            Left            =   1446
            MaxLength       =   10
            TabIndex        =   1
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label LblCMC7 
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "CMC-7"
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
            Left            =   204
            TabIndex        =   26
            Top             =   240
            Width           =   636
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valores"
         Height          =   636
         Left            =   5352
         TabIndex        =   22
         Top             =   2484
         Width           =   3228
         Begin CURRENCYEDITLib.CurrencyEdit TxtCheques 
            Height          =   372
            Left            =   1044
            TabIndex        =   8
            Top             =   204
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
         Begin CURRENCYEDITLib.CurrencyEdit TxtTotal 
            Height          =   372
            Left            =   1044
            TabIndex        =   9
            Top             =   636
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
         Begin VB.Label label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   480
            TabIndex        =   24
            Top             =   720
            Width           =   444
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Total"
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
            Left            =   96
            TabIndex        =   23
            Top             =   288
            Width           =   912
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados"
         Height          =   1092
         Left            =   132
         TabIndex        =   3
         Top             =   2028
         Width           =   5148
         Begin VB.TextBox txtAgencia 
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
            Left            =   1308
            MaxLength       =   4
            TabIndex        =   5
            Top             =   588
            Width           =   624
         End
         Begin VB.TextBox txtIdentificado 
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
            Left            =   180
            MaxLength       =   6
            TabIndex        =   4
            Top             =   588
            Width           =   876
         End
         Begin VB.TextBox txtConta 
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
            Left            =   2136
            MaxLength       =   7
            TabIndex        =   6
            Top             =   588
            Width           =   936
         End
         Begin VB.ComboBox CboTipoConta 
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
            Height          =   336
            ItemData        =   "Deposito.frx":030A
            Left            =   3204
            List            =   "Deposito.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   588
            Width           =   1848
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
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
            Left            =   1296
            TabIndex        =   21
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
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
            Left            =   2172
            TabIndex        =   20
            Top             =   300
            Width           =   528
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificado"
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
            Left            =   132
            TabIndex        =   19
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Conta"
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
            Left            =   3204
            TabIndex        =   18
            Top             =   300
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   2856
         Picture         =   "Deposito.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   216
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   3672
         Picture         =   "Deposito.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   216
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   4488
         Picture         =   "Deposito.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   216
         Width           =   804
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   696
         Left            =   5304
         Picture         =   "Deposito.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   216
         Width           =   804
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   6108
         Picture         =   "Deposito.frx":0F36
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   216
         Width           =   804
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   696
         Left            =   7728
         Picture         =   "Deposito.frx":1240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   216
         Width           =   816
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   180
         Picture         =   "Deposito.frx":154A
         Top             =   288
         Width           =   384
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Depósitos"
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
         Left            =   708
         TabIndex        =   17
         Top             =   408
         Width           =   912
      End
   End
End
Attribute VB_Name = "Deposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Vairáveis do RDO
Private qryGetDepositoDuplicado As rdoQuery
Private qryAtualizaDocumentoExcluido As rdoQuery
Private qryAtualizaDeposito As rdoQuery
Private qryGetDeposito As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Sub AjustesIniciais()

  On Error GoTo ERRO_AJUSTESINICIAIS

  'Preencher o Combo de Tipos de Conta
  CboTipoConta.AddItem ""

  CboTipoConta.AddItem "0 - Corrente"
  CboTipoConta.ItemData(CboTipoConta.NewIndex) = 2

  CboTipoConta.AddItem "9 - Poupança"
  CboTipoConta.ItemData(CboTipoConta.NewIndex) = 3

  'Setar o Tipo da Conta para 'CONTA CORRENTE' como default
  CboTipoConta.ListIndex = 0

  'Setar os Objetos RDOQuery
  Set qryGetDepositoDuplicado = Geral.Banco.CreateQuery("", "{? = call GetDepositoDuplicado (?,?,?)}")
  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaDeposito = Geral.Banco.CreateQuery("", "{? = call AtualizaDeposito (?,?,?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")

  Exit Sub

ERRO_AJUSTESINICIAIS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do Depósito.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CamposOK() As Boolean

  'Primeiro Campo do CMC7
  If Len(Trim(txtCMC71.Text)) = 0 Then
    MsgBox "Informe o Primeiro Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC71.SetFocus
    Exit Function
  End If

  'Segundo Campo do CMC7
  If Len(Trim(txtCMC72.Text)) = 0 Then
    MsgBox "Informe o Segundo Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC72.SetFocus
    Exit Function
  End If

  'Terceiro Campo do CMC7
  If Len(Trim(txtCMC73.Text)) = 0 Then
    MsgBox "Informe o Terceiro Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC73.SetFocus
    Exit Function
  End If

  'Agencia
  If Len(Trim(txtAgencia.Text)) = 0 Then
    MsgBox "Informe o Código da Agencia.", vbInformation, App.Title
    CamposOK = False
    txtAgencia.SetFocus
    Exit Function
  End If

  'Conta
  If Len(Trim(txtConta.Text)) = 0 Then
    MsgBox "Informe o Número da Conta.", vbInformation, App.Title
    CamposOK = False
    txtConta.SetFocus
    Exit Function
  End If

  'Tipo da Conta
  If CboTipoConta.ListIndex < 1 Then
    MsgBox "Informe o tipo da conta.", vbInformation + vbOKOnly, App.Title
    CamposOK = False
    CboTipoConta.SetFocus
    Exit Function
  End If

  'Valor dos Cheques
  If Len(Trim(TxtCheques.Text)) = 0 Or Val(TxtCheques.Text) = 0 Then
    MsgBox "Informe o Valor dos Cheques.", vbInformation, App.Title
    CamposOK = False
    TxtCheques.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Sub PesquisaDeposito()

  On Error GoTo ERRO_PESQUISADEPOSITO

  Dim sSql As String
  Dim RsDeposito As rdoResultset
  Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, svalor As String

  'Pesquisar o Deposito Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetDeposito = Geral.Banco.CreateQuery("", "{call GetDeposito (" & sSql & ")}")

  Set RsDeposito = qryGetDeposito.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsDeposito.EOF Then
    'Encontrou o Deposito -> Preencher os campos
    txtIdentificado.Text = RsDeposito!Identificado
    txtAgencia.Text = Format(RsDeposito!Agencia, "0000")
    txtConta.Text = RsDeposito!Conta
    CboTipoConta.ListIndex = Val(RsDeposito!TipoConta)
    TxtCheques.Text = RsDeposito!Cheque * 100
    TxtTotal.Text = RsDeposito!Valor * 100
  End If

  If Len(Trim(Geral.Documento.Leitura)) <> 0 Then
    txtCMC71.Text = Mid(Geral.Documento.Leitura, 1, 8)
    txtCMC72.Text = Mid(Geral.Documento.Leitura, 9, 10)
    txtCMC73.Text = Mid(Geral.Documento.Leitura, 19, 12)

    If Mid(Geral.Documento.Leitura, 1, 3) = "409" And Mid(Geral.Documento.Leitura, 9, 3) = "999" Then
      'Verifica se posiciona em CMC7 ou Valor
      If Not TratarCamposCMC7(Geral.Documento.Leitura, sCampo1, sCampo2, sCampo3, svalor) Then
        txtCMC71.SetFocus
      Else
        txtIdentificado.SetFocus
      End If
    Else
      txtCMC71.SetFocus
    End If
  Else
    txtCMC71.SetFocus
  End If

  Exit Sub

ERRO_PESQUISADEPOSITO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados do Depósito.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function SalvaDeposito() As Boolean

    On Error GoTo ERRO_SALVADEPOSITO

    Dim vCMC7           As String
    Dim sCampo1         As String
    Dim sCampo2         As String
    Dim sCampo3         As String
    Dim svalor          As String
    
    Dim sTipo           As String
    Dim sTamanho        As Integer
    Dim strEncripta     As String
    
    SalvaDeposito = False

    'Verificar se todos os campos estão preenchidos
    If CamposOK Then
        'Verificar se as tres primeiras posições do primeiro campo do CMC7 devem ser iguais à : 409
        If Mid(txtCMC71.Text, 1, 3) <> "409" Then
            MsgBox "O CMC7 do Depósito não é Válido.", vbInformation, App.Title
            txtCMC71.SetFocus
            Exit Function
        End If

        'Verificar se as tres primeiras posições do segundo campo do CMC7 devem ser iguais à : 999
        If Mid(txtCMC72.Text, 1, 3) <> "999" Then
            MsgBox "O CMC7 do Depósito não é Válido.", vbInformation, App.Title
            txtCMC72.SetFocus
            Exit Function
        End If

        'Validar Agencia e Conta
        sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
        If Val(txtAgencia.Text & txtConta.Text) <> 0 Then
            If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
                MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
                txtAgencia.SetFocus
                Exit Function
            End If
        Else
            MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
            txtAgencia.SetFocus
            Exit Function
        End If

        'Validar CMC7
        vCMC7 = txtCMC71.Text & txtCMC72.Text & txtCMC73.Text

        If Not TratarCamposCMC7(vCMC7, sCampo1, sCampo2, sCampo3, svalor) Then
            MsgBox "CMC7 Inválido.", vbInformation, App.Title
            'Verificar qual campo está zerado e posicionar o cursor
            If Val(sCampo1) = 0 Then
                If txtCMC71.Visible = True Then
                    txtCMC71.SetFocus
                End If
                Exit Function
            End If
        
            If Val(sCampo2) = 0 Then
                If txtCMC72.Visible = True Then
                    txtCMC72.SetFocus
                End If
                Exit Function
            End If
        
            If Val(sCampo3) = 0 Then
                If txtCMC73.Visible = True Then
                    txtCMC73.SetFocus
                End If
                Exit Function
            End If
        End If
    
        'Verificar se o Documento pertence à outro Tipo
        If Geral.Documento.TipoDocto <> 2 And Geral.Documento.TipoDocto <> 3 And Geral.Documento.TipoDocto <> 0 Then
            With qryRemoveTipoDocumento
                .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
                .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
                .Execute
            End With
        End If
    
        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(CboTipoConta.ItemData(CboTipoConta.ListIndex), CStr(Val(txtConta.Text)))
        If strEncripta = "" Then GoTo ERRO_SALVADEPOSITO
        
        'Atualizar / Inserir Deposito (AtualizaDeposito)
        With qryAtualizaDeposito
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.DataProcessamento                       'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto                       'IdDocto
            .rdoParameters(3) = vCMC7                                         'CMC7
            .rdoParameters(4) = Val(txtIdentificado.Text)                     'Identificado
            .rdoParameters(5) = Format(txtAgencia.Text, "0000")               'Agencia
            .rdoParameters(6) = txtConta.Text                                 'Conta
            .rdoParameters(7) = Val(CboTipoConta.ListIndex)                   'TipoConta
            .rdoParameters(8) = Val(TxtTotal.Text) / 100                      'Valor
            .rdoParameters(9) = CboTipoConta.ItemData(CboTipoConta.ListIndex) 'TipoDocto
            .rdoParameters(10) = strEncripta                                  'Autenticacao digital
            .Execute
        End With
    
        If qryAtualizaDeposito(0).Value = 2 Then
            Geral.Documento.Status = "D"
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Envia para confirmação somente se o usuario for terceiro e o docto ñ for duplicidade'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If GrupoUsuario(Geral.Usuario, eG_TERCEIRO) And Geral.Documento.Status <> "D" Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Não faz nada caso não conseguiu atualizar o status do documento'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not ConfirmaAgConta(Geral.Documento.IdDocto) Then
                MsgBox "Não foi possível enviar este documento para confirmação de Agência e Conta.", vbCritical
                Exit Function
            End If
            Geral.Documento.Status = "L"
        End If
    
        SalvaDeposito = True
    
        'Atualizar o Controle Global
        Geral.Documento.ValorTotal = Val(TxtTotal.Text) / 100
        Geral.Documento.Leitura = vCMC7
        Geral.Documento.TipoDocto = CboTipoConta.ItemData(CboTipoConta.ListIndex)
    End If

    Exit Function

ERRO_SALVADEPOSITO:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar Dados do Depósito.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function
Private Sub CboTipoConta_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKey0 Or KeyAscii = vbKey9 Then
    SendKeys "{TAB}"
  Else
    CboTipoConta.ListIndex = -1
    KeyAscii = 0
  End If
End Sub
Private Sub cmdConfirmar_Click()

  'Atualizar o Campo 'TOTAL'
  TxtTotal.Text = TxtCheques.Text

  'Valida preenchimento máximo de CMC7
  If VerificaPreenchimentoCMC7(Me) = False Then Exit Sub

  If SalvaDeposito Then
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


Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub

Private Sub Form_Activate()

  Call AjustesIniciais
  
  Call PesquisaDeposito

  Screen.MousePointer = vbDefault
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

  'Fechar as conexões (RDOQUERY)
  Set qryGetDepositoDuplicado = Nothing
  Set qryAtualizaDocumentoExcluido = Nothing
  Set qryAtualizaDeposito = Nothing
End Sub

Private Sub txtAgencia_Change()
  If Len(Trim(txtAgencia.Text)) = txtAgencia.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtAgencia_GotFocus()
  txtAgencia.SelStart = 0
  txtAgencia.SelLength = txtAgencia.MaxLength
End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtAgencia_LostFocus()
    
    If Len(Trim(txtAgencia.Text)) > 0 Then
      'Valida Agencia
      If Not IsNumeric(txtAgencia.Text) Then
        MsgBox "Número de Agência inválido, Redigite.", vbInformation, App.Title
        txtAgencia.Text = ""
        txtAgencia.SetFocus
        Exit Sub
      End If
    End If
      
End Sub

Private Sub txtCheques_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    'Atualizar o Campo 'TOTAL'
    TxtTotal.Text = TxtCheques.Text

    Call cmdConfirmar_Click
  End If
End Sub

Private Sub TxtCheques_LostFocus()

  TxtTotal.Text = TxtCheques.Text
End Sub
Private Sub txtCMC71_Change()

  If Len(Trim(txtCMC71.Text)) = txtCMC71.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub


Private Sub txtCMC71_GotFocus()

  txtCMC71.SelStart = 0
  txtCMC71.SelLength = txtCMC71.MaxLength
End Sub

Private Sub txtCMC71_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtCMC72_Change()

  If Len(Trim(txtCMC72.Text)) = txtCMC72.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub


Private Sub txtCMC72_GotFocus()

  txtCMC72.SelStart = 0
  txtCMC72.SelLength = txtCMC72.MaxLength
End Sub

Private Sub txtCMC72_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtCMC73_Change()

  If Len(Trim(txtCMC73.Text)) = txtCMC73.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtCMC73_GotFocus()

  txtCMC73.SelStart = 0
  txtCMC73.SelLength = txtCMC73.MaxLength
End Sub

Private Sub txtCMC73_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtConta_Change()
  If Len(Trim(txtConta.Text)) = txtConta.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtConta_GotFocus()
  txtConta.SelStart = 0
  txtConta.SelLength = txtConta.MaxLength
End Sub
Private Sub txtConta_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtConta_LostFocus()

   Dim sTamanho As String

    If Len(Trim(txtAgencia.Text)) > 0 And Len(Trim(txtConta.Text)) > 0 Then
        'Valida Conta
        If Not IsNumeric(txtConta.Text) Then
            MsgBox "Número de Conta inválido, Redigite.", vbInformation, App.Title
            txtConta.Text = ""
            txtConta.SetFocus
            Exit Sub
        End If

        'Valida Agencia & Conta
        If Val(txtAgencia.Text) <> 0 And Val(txtConta.Text) <> 0 Then
            sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
            If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
                MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
                txtConta.SelStart = 0
                txtConta.SelLength = Len(txtConta.Text)
                txtAgencia.SetFocus
            End If
        Else
            MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
            txtAgencia.SetFocus
        End If
    End If
End Sub
Private Sub txtIdentificado_Change()

  If Len(Trim(txtIdentificado.Text)) = txtIdentificado.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtIdentificado_GotFocus()

  txtIdentificado.SelStart = 0
  txtIdentificado.SelLength = txtIdentificado.MaxLength
End Sub


Private Sub txtIdentificado_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtTotal_GotFocus()
 
  TxtTotal.SelStart = 0
  TxtTotal.SelLength = Len(TxtTotal.Text)
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    'Atualizar o Campo 'TOTAL'
    TxtTotal.Text = TxtCheques.Text

    Call cmdConfirmar_Click
  End If
End Sub
