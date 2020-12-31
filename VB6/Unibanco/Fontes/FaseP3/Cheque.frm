VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form Cheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Cheques"
   ClientHeight    =   1992
   ClientLeft      =   1272
   ClientTop       =   1284
   ClientWidth     =   9408
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1992
   ScaleWidth      =   9408
   Begin VB.Frame Frame1 
      Height          =   1944
      Left            =   36
      TabIndex        =   13
      Top             =   -24
      Width           =   9288
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   720
         Left            =   7512
         Picture         =   "Cheque.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   192
         Width           =   816
      End
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   372
         Left            =   6672
         TabIndex        =   12
         Top             =   1404
         Width           =   2484
         _Version        =   65537
         _ExtentX        =   4382
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
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   720
         Left            =   8328
         Picture         =   "Cheque.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   192
         Width           =   816
      End
      Begin VB.CommandButton cmdLinha1 
         Caption         =   "Linha 1"
         Height          =   720
         Left            =   6696
         Picture         =   "Cheque.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton CmdCMC7 
         Caption         =   "CMC7"
         Height          =   720
         Left            =   5880
         Picture         =   "Cheque.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   720
         Left            =   5064
         Picture         =   "Cheque.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   720
         Left            =   4248
         Picture         =   "Cheque.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   720
         Left            =   3432
         Picture         =   "Cheque.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   720
         Left            =   2616
         Picture         =   "Cheque.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   720
         Left            =   1800
         Picture         =   "Cheque.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   192
         Width           =   804
      End
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
         Left            =   1376
         MaxLength       =   4
         TabIndex        =   2
         Top             =   2496
         Visible         =   0   'False
         Width           =   696
      End
      Begin VB.TextBox txtComp 
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
         Left            =   192
         MaxLength       =   3
         TabIndex        =   0
         Top             =   2496
         Visible         =   0   'False
         Width           =   456
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
         Left            =   2644
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2496
         Visible         =   0   'False
         Width           =   1308
      End
      Begin VB.TextBox txtDV3 
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
         Left            =   5488
         MaxLength       =   1
         TabIndex        =   7
         Top             =   2496
         Visible         =   0   'False
         Width           =   336
      End
      Begin VB.TextBox txtTipo 
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
         Left            =   5916
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2496
         Visible         =   0   'False
         Width           =   348
      End
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
         Left            =   192
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1416
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
         Left            =   3324
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1416
         Width           =   1536
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
         Left            =   1656
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1416
         Width           =   1296
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
         Left            =   784
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2496
         Visible         =   0   'False
         Width           =   456
      End
      Begin VB.TextBox txtDV1 
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
         Left            =   2208
         MaxLength       =   1
         TabIndex        =   3
         Top             =   2496
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtDV2 
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
         Left            =   4088
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2496
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtCheque 
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
         Left            =   4524
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2496
         Visible         =   0   'False
         Width           =   828
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   192
         TabIndex        =   26
         Top             =   1164
         Visible         =   0   'False
         Width           =   432
      End
      Begin VB.Label Lbltipo 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   5904
         TabIndex        =   32
         Top             =   1164
         Visible         =   0   'False
         Width           =   396
      End
      Begin VB.Label lblDV3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "C3"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   5544
         TabIndex        =   31
         Top             =   1164
         Visible         =   0   'False
         Width           =   228
      End
      Begin VB.Label lblcheque 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cheq"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4560
         TabIndex        =   30
         Top             =   1164
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblDv2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "C2"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4176
         TabIndex        =   29
         Top             =   1164
         Visible         =   0   'False
         Width           =   228
      End
      Begin VB.Label lblConta 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   2568
         TabIndex        =   28
         Top             =   1164
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label lblDV1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "C1"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   2256
         TabIndex        =   27
         Top             =   1164
         Visible         =   0   'False
         Width           =   228
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   1416
         TabIndex        =   25
         Top             =   1164
         Visible         =   0   'False
         Width           =   696
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   816
         TabIndex        =   24
         Top             =   1164
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
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
         Left            =   6660
         TabIndex        =   23
         Top             =   1164
         Width           =   468
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cheques"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   600
         TabIndex        =   21
         Top             =   408
         Width           =   648
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   60
         Picture         =   "Cheque.frx":1C92
         Top             =   288
         Width           =   384
      End
      Begin VB.Label LblCMC7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "CMC-7"
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
         Left            =   204
         TabIndex        =   22
         Top             =   1116
         Width           =   624
      End
   End
End
Attribute VB_Name = "Cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variáveis do RDO
'Private qryGetChequeDuplicado As rdoQuery
Private qryAtualizaCheque As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
'Private qryAtualizaDocumentoExcluido As rdoQuery

'Variáveis de Controle
Private teclou As Boolean
Private bFormating As Boolean
Private vCMC7 As String
Private Linha1 As Boolean
Private mForm As Form
Public Alterou As Boolean
Sub PreencheCampos()

Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, svalor As String

  On Error GoTo ERRO_PREENCHECMC7

  'Preencher o Campo VALOR
  TxtValor.Text = Geral.Documento.ValorTotal

  'Preencher os campos CMC7 e posicionar o cursor no campo de VALOR
  If Len(Trim(Geral.Documento.Leitura)) > 0 Then
    'Preencher os campos de CMC7 com o campo LEITURA
    txtCMC71.Text = Left(Geral.Documento.Leitura, 8)
    txtCMC72.Text = Mid(Geral.Documento.Leitura, 9, 10)
    txtCMC73.Text = Mid(Geral.Documento.Leitura, 19, 12)

    If Mid(Geral.Documento.Leitura, 9, 3) <> "256" Then
      'Verifica se posiciona em CMC7 ou Valor
      If Not TratarCamposCMC7(Geral.Documento.Leitura, sCampo1, sCampo2, sCampo3, svalor) Then
          txtCMC71.SetFocus
      Else
          TxtValor.SetFocus
      End If
    Else
      txtCMC71.SetFocus
    End If
  Else
    txtCMC71.SetFocus
  End If

'  Set qryGetChequeDuplicado = Geral.Banco.CreateQuery("", "{? = call GetChequeDuplicado (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryAtualizaCheque = Geral.Banco.CreateQuery("", "{? = call AtualizaCheque (?,?,?,?,?,?)}")
'  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoExcluido (?,?,?,?,?)}")

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PREENCHECMC7:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Preencher CMC7.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub SalvaCheque()

    On Error GoTo ERRO_SALVACHEQUE

    Dim Valor As Currency
    Dim RsCheque As rdoResultset
    Dim sSql As String
    Dim TipoDocto As String
    Dim strEncripta   As String
  
    'Definir o Tipo do Documento
    If Left(vCMC7, 3) = "409" Or Left(vCMC7, 3) = "230" Then
        'Cheque Unibanco / Bandeirantes
        TipoDocto = "5"
    Else
        'Cheque Terceiro
        TipoDocto = "6"
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 5 And Geral.Documento.TipoDocto <> 6 And Geral.Documento.TipoDocto <> 0 Then
        With qryRemoveTipoDocumento
            .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
            .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
            .Execute
        End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(TipoDocto, vCMC7)
    If strEncripta = "" Then GoTo ERRO_SALVACHEQUE

    'Inserir / Atualizar registro na tabela 'CHEQUE'
    With qryAtualizaCheque
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = vCMC7                       'CMC7
        .rdoParameters(4) = Val(TxtValor.Text) / 100    'Valor
        .rdoParameters(5) = TipoDocto                   'TipoDocto
        .rdoParameters(6) = strEncripta                 'Autenticacao digital
        .Execute
    End With

    If qryAtualizaCheque(0).Value = 2 Then
        'Documento Duplicado
        Geral.Documento.Status = "D"
    End If

    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = vCMC7
    Geral.Documento.TipoDocto = TipoDocto

    Alterou = True
    Me.Hide

    Exit Sub

ERRO_SALVACHEQUE:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Salvar Dados do Cheque.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Function VerificarDV3() As Boolean

  On Error GoTo ERRO_VERIFICADV3

  Dim Campo3 As String
  Dim Ret3 As String

  VerificarDV3 = True

  If Len(txtCheque) = 0 Then
      MsgBox "Digite o campo Número Cheque.", vbExclamation + vbOKOnly, "Atenção"
      txtCheque.SetFocus
      VerificarDV3 = False
      Exit Function
  Else
      txtCheque = Format(txtCheque, "000000")
  End If

  If Len(txtDV3) = 0 Then
      MsgBox "Digite o campo Digito Verificador 3.", vbExclamation + vbOKOnly, "Atenção"
      txtDV3.SetFocus
      VerificarDV3 = False
      Exit Function
  End If

  Campo3 = txtCheque & txtDV3

  Ret3 = Modulo11(Campo3)

  If Ret3 = False Then
      MsgBox "Terceiro lote da linha1 não confere. Verifique.", vbExclamation + vbOKOnly, Caption
      txtCheque = ""
      txtDV3 = ""
      txtCheque.SetFocus
      VerificarDV3 = False
      Exit Function
  End If

  Exit Function

ERRO_VERIFICADV3:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar terceiro Lote.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificarDV1() As Boolean

  On Error GoTo ERRO_VERIFICARDV1

  Dim Campo1 As String
  Dim ret1 As String

  VerificarDV1 = True

  'Valida campo compensação
  If Len(txtComp) = 0 Then
      MsgBox "Digite o campo Comp.", vbExclamation + vbOKOnly, "Atenção"
      txtComp.SetFocus
      VerificarDV1 = False
      Exit Function
  Else
      txtComp = Format(txtComp, "000")
  End If

  'Valida o campo Banco
  If Len(txtBanco) = 0 Then
      MsgBox "Digite o campo Banco.", vbExclamation + vbOKOnly, "Atenção"
      txtBanco.SetFocus
      VerificarDV1 = False
      Exit Function
  Else
      txtBanco = Format(txtBanco, "000")
  End If

  'Valida o campo Agencia
  If Len(txtAgencia) = 0 Then
      MsgBox "Digite o campo Agencia.", vbExclamation + vbOKOnly, "Atenção"
      txtAgencia.SetFocus
      VerificarDV1 = False
      Exit Function
  Else
      txtAgencia = Format(txtAgencia, "0000")
  End If

  'Valida o DV1 linha1
  If Len(txtDV1) = 0 Then
      MsgBox "Digite o campo Digito Verificador 1.", vbExclamation + vbOKOnly, "Atenção"
      txtDV1.SetFocus
      VerificarDV1 = False
      Exit Function
  End If

  Campo1 = txtComp & txtBanco & txtAgencia & txtDV1

  ret1 = Modulo11(Campo1)

  If ret1 = False Then
      MsgBox "Primeiro lote da linha 1 não confere. Verifique.", vbExclamation + vbOKOnly, Caption
      VerificarDV1 = False
      txtComp = ""
      txtBanco = ""
      txtAgencia = ""
      txtDV1 = ""
      txtComp.SetFocus
      Exit Function
  End If

  Exit Function

ERRO_VERIFICARDV1:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar primeiro Lote.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificarDV2() As Boolean

  On Error GoTo ERRO_VERIFICADV2

  Dim Campo2 As String
  Dim ret2 As String

  VerificarDV2 = True

  If Len(txtConta) = 0 Then
      MsgBox "Digite o campo Conta.", vbExclamation + vbOKOnly, "Atenção"
      txtConta.SetFocus
      VerificarDV2 = False
      Exit Function
  Else
      'alteração 160300 - by Madriana
      txtConta = Format(txtConta, "0000000000")
  End If

  If Len(txtDV2) = 0 Then
      MsgBox "Digite o campo Digito Verificador 2.", vbExclamation + vbOKOnly, "Atenção"
      txtDV2.SetFocus
      VerificarDV2 = False
      Exit Function
  End If

  Campo2 = txtConta & txtDV2

  ret2 = Modulo11(Campo2)

  If ret2 = False Then
      MsgBox "Segundo lote da linha1 não confere. Verifique.", vbExclamation + vbOKOnly, Caption
      txtConta = ""
      txtDV2 = ""
      txtConta.SetFocus
      VerificarDV2 = False
      Exit Function
  End If

  Exit Function

ERRO_VERIFICADV2:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar segundo Lote.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub DigitaCMC7()

  On Error GoTo ERRO_DigitaCMC7

  'Posicionar os objetos para digitação de CMC7
  txtCMC71.Top = TxtValor.Top
  txtCMC72.Top = TxtValor.Top
  txtCMC73.Top = TxtValor.Top
  LblCMC7.Top = TxtValor.Top - 250

  'Deixar Objetos CMC7 visíveis e Linha1 invisíveis
  txtCMC71.Visible = True
  txtCMC72.Visible = True
  txtCMC73.Visible = True
  LblCMC7.Visible = True

  txtComp.Visible = False
  txtBanco.Visible = False
  txtAgencia.Visible = False
  txtDV1.Visible = False
  txtConta.Visible = False
  txtDV2.Visible = False
  txtCheque.Visible = False
  txtDV3.Visible = False
  txtTipo.Visible = False
  lblComp.Visible = False
  lblBanco.Visible = False
  lblAgencia.Visible = False
  lblDV1.Visible = False
  lblConta.Visible = False
  lblDv2.Visible = False
  lblcheque.Visible = False
  lblDV3.Visible = False
  Lbltipo.Visible = False

  'Setar o foco para o primeiro campo de CMC7
  txtCMC71.TabIndex = 0
  txtCMC71.SetFocus
  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_DigitaCMC7:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao preparar campos para Digitação de CMC7.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub DigitaLinha1()

  On Error GoTo ERRO_DIGITALINHA1

  'Posicionar os objetos para digitação da Linha 1
  txtComp.Top = TxtValor.Top
  txtBanco.Top = TxtValor.Top
  txtAgencia.Top = TxtValor.Top
  txtDV1.Top = TxtValor.Top
  txtConta.Top = TxtValor.Top
  txtDV2.Top = TxtValor.Top
  txtCheque.Top = TxtValor.Top
  txtDV3.Top = TxtValor.Top
  txtTipo.Top = TxtValor.Top

  lblComp.Top = TxtValor.Top - 250
  lblBanco.Top = TxtValor.Top - 250
  lblAgencia.Top = TxtValor.Top - 250
  lblDV1.Top = TxtValor.Top - 250
  lblConta.Top = TxtValor.Top - 250
  lblDv2.Top = TxtValor.Top - 250
  lblcheque.Top = TxtValor.Top - 250
  lblDV3.Top = TxtValor.Top - 250
  Lbltipo.Top = TxtValor.Top - 250

  'Deixar Objetos da Linha 1 visíveis e CMC7 invisíveis
  txtComp.Visible = True
  txtBanco.Visible = True
  txtAgencia.Visible = True
  txtDV1.Visible = True
  txtConta.Visible = True
  txtDV2.Visible = True
  txtCheque.Visible = True
  txtDV3.Visible = True
  txtTipo.Visible = True
  lblComp.Visible = True
  lblBanco.Visible = True
  lblAgencia.Visible = True
  lblDV1.Visible = True
  lblConta.Visible = True
  lblDv2.Visible = True
  lblcheque.Visible = True
  lblDV3.Visible = True
  Lbltipo.Visible = True

  txtCMC71.Visible = False
  txtCMC72.Visible = False
  txtCMC73.Visible = False
  LblCMC7.Visible = False

  'Setar o foco para o primeiro campo de CMC7
  txtComp.SetFocus

  Exit Sub

ERRO_DIGITALINHA1:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao preparar campos para Digitação da Linha1.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Function ValidaLinha1()

  On Error GoTo ERRO_VALIDALINHA1

  ValidaLinha1 = True

  'Verifica 1º lote da linha1
  If VerificarDV1 = False Then
    ValidaLinha1 = False
    Exit Function
  End If

  'Verifica 2º lote da linha1
  If VerificarDV2 = False Then
    ValidaLinha1 = False
    Exit Function
  End If

  'Verifica 3º lote da linha1
  If VerificarDV3 = False Then
    ValidaLinha1 = False
    Exit Function
  End If

  Call CalculaDVlinha1

  Exit Function

ERRO_VALIDALINHA1:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Linha1.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdCMC7_Click()

  Call DigitaCMC7
  Linha1 = False
End Sub

Private Sub cmdConfirmar_Click()

    On Error GoTo ERRO_CONFIRMAR

    Dim sCampo1     As String
    Dim sCampo2     As String
    Dim sCampo3     As String
    Dim svalor      As String
    Dim sTamanho    As Integer

    'Valida preenchimento máximo de CMC7
     If Linha1 = False Then
        If VerificaPreenchimentoCMC7(Me) = False Then Exit Sub
     End If

    'Verificar se foi informado CMC7 ou Linha1
    If Linha1 Then
        'Verificar se foi informado um tipo válido
        If Val(txtTipo.Text) <> 5 And Val(txtTipo.Text) <> 6 And Val(txtTipo.Text) <> 8 And Val(txtTipo.Text) <> 9 Then
            MsgBox "Tipificação Inválida.", vbInformation, App.Title
            txtTipo.SetFocus
            Exit Sub
        End If

        'Formatar Numero da Conta
        txtConta.Text = Format(txtConta.Text, "0000000000")

        'Validar o Código do Banco
        If Len(Trim(txtBanco.Text)) <> 0 Then
            If Not ValidaCodigoBanco(txtBanco.Text) Then
                MsgBox "Código de Banco não participante do Sistema de Compensação.", vbInformation, App.Title
                txtBanco.SetFocus
                Exit Sub
            End If
        End If

        'Validar Primeiro Lote da Linha1
        If Val(txtComp.Text) <> 0 And Val(txtBanco.Text) <> 0 And Len(Trim(txtDV1.Text)) > 0 Then
            If Not Modulo11(txtComp.Text & txtBanco.Text & txtAgencia.Text & txtDV1.Text) Then
                MsgBox "Primeiro Lote da Linha 1 é Inválido.", vbInformation, App.Title
                txtComp.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Primeiro Lote da Linha 1 é Inválido.", vbInformation, App.Title
            txtComp.SetFocus
            Exit Sub
        End If

        'Validar Segundo Lote da Linha1
        If Val(txtConta.Text) <> 0 And Len(Trim(txtDV2.Text)) > 0 Then
            If Not Modulo11(txtConta.Text & txtDV2.Text) Then
                MsgBox "Segundo Lote da Linha 1 é Inválido.", vbInformation, App.Title
                txtConta.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Segundo Lote da Linha 1 é Inválido.", vbInformation, App.Title
            txtConta.SetFocus
            Exit Sub
        End If

        'Validar Terceiro Lote da Linha1
        If Val(txtCheque.Text) <> 0 And Len(Trim(txtDV3.Text)) > 0 Then
            If Not Modulo11(txtCheque.Text & txtDV3.Text) Then
                MsgBox "Terceiro Lote da Linha 1 é Inválido.", vbInformation, App.Title
                txtCheque.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Terceiro Lote da Linha 1 é Inválido.", vbInformation, App.Title
            txtCheque.SetFocus
            Exit Sub
        End If

        'Transformar Linha1 em CMC7
        Call CalculaDVlinha1
    End If

    'Concatenar CMC7 com os campos do CMC7
    vCMC7 = txtCMC71.Text & txtCMC72.Text & txtCMC73.Text

    'Verificar se a tipificação é válida
    If Val(Mid(vCMC7, 18, 1)) <> 5 And Val(Mid(vCMC7, 18, 1)) <> 6 And Val(Mid(vCMC7, 18, 1)) <> 8 And Val(Mid(vCMC7, 18, 1)) <> 9 Then
        MsgBox "Tipificação Inválida.", vbInformation, App.Title
        If txtTipo.Visible = True Then
            txtTipo.SetFocus
        Else
            txtCMC72.SetFocus
        End If
        Exit Sub
    End If

    'Validar o Código do Banco
    If Len(Trim(Left(vCMC7, 3))) <> 0 Then
        If Not ValidaCodigoBanco(Left(vCMC7, 3)) Then
            MsgBox "Código de Banco não participante do Sistema de Compensação.", vbInformation, App.Title
            If txtBanco.Visible = True Then
                txtBanco.SetFocus
            Else
                txtCMC71.SetFocus
            End If
            Exit Sub
        End If
    End If

    'Verificar se o CMC7 digitado é um CMC7 válido para depósito , ADCC , Capa Malote ou Capa OCT
    If (Left(txtCMC71.Text, 3) = "409" And Left(txtCMC72.Text, 3) = "999") Or (Left(txtCMC71.Text, 3) = "409" And Left(txtCMC72.Text, 3) = "256") Or (Left(txtCMC71.Text, 3) = "409" And Left(txtCMC72.Text, 3) = "600") Or (Left(txtCMC71.Text, 3) = "409" And Left(txtCMC72.Text, 3) = "592") Then
        MsgBox "Este CMC7 não é válido para Cheque.", vbInformation, App.Title
        txtCMC71.SetFocus
        Exit Sub
    End If

    'Verificar se a agencia e conta do CMC7 são validos para cheques do Unibanco
    If Left(txtCMC71.Text, 3) = "409" Then
        sTamanho = Len(Mid(vCMC7, 4, 4) & Mid(vCMC7, 23, 7))
        If Val(Mid(vCMC7, 4, 4) & Mid(vCMC7, 23, 7)) <> 0 Then
            If Not Modulo10(Mid(vCMC7, 4, 4) & Mid(vCMC7, 23, 7), sTamanho) Or Val(Mid(vCMC7, 12, 6)) = 0 Then
                MsgBox "Agência, Nro. Cheque e/ou Conta Inválidos!", vbInformation, App.Title
                If txtCMC71.Visible = True Then txtCMC71.SetFocus
                Exit Sub
            End If
        End If
    Else
        'Verifica se Banco, agência, Nro. Cheque e Conta outros Bancos são válidos
        If Val(Left(vCMC7, 3)) = 0 Or Val(Mid(vCMC7, 4, 4)) = 0 Or _
            Val(Mid(vCMC7, 12, 6)) = 0 Or Val(Mid(vCMC7, 20, 10)) = 0 Then
            MsgBox "Banco, Agência, Nro. Cheque e/ou Conta Inválidos!", vbInformation, App.Title
            If txtCMC71.Visible = True Then txtCMC71.SetFocus
            Exit Sub
        End If
    End If

    If Not TratarCamposCMC7(vCMC7, sCampo1, sCampo2, sCampo3, svalor) Then
        MsgBox "CMC7 Inválido.", vbInformation, App.Title
        'Verificar qual campo está zerado e posicionar o cursor
        If Val(sCampo1) = 0 Then
            If txtCMC71.Visible = True Then
                txtCMC71.SetFocus
            Else
                txtBanco.SetFocus
            End If
            Exit Sub
        End If

        If Val(sCampo2) = 0 Then
            If txtCMC72.Visible = True Then
                txtCMC72.SetFocus
            End If
            Exit Sub
        End If

        If Val(sCampo3) = 0 Then
            If txtCMC73.Visible = True Then
                txtCMC73.SetFocus
            End If
            Exit Sub
        End If
    End If

    'Verificar se foi informado o Valor do Cheque
    If Val(TxtValor.Text) = 0 Then
        MsgBox "Informe o Valor do Cheque.", vbInformation, App.Title
        TxtValor.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Call SalvaCheque
    Screen.MousePointer = vbDefault

    Exit Sub

ERRO_CONFIRMAR:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar Dados do Cheque.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Private Sub cmdFrenteVerso_Click()

  mForm.cmdFrenteVerso_Click
End Sub

Private Sub cmdInverteCor_Click()

  mForm.cmdInverteCor_Click
End Sub
Private Sub cmdLinha1_Click()
  Call DigitaLinha1
  Linha1 = True
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
  Linha1 = False
  Call PreencheCampos
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
      Call mForm.Form_KeyUp(KeyCode, Shift)
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = 0 Then Alterou = False
  
'  Set qryGetChequeDuplicado = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryAtualizaCheque = Nothing
'  Set qryAtualizaDocumentoExcluido = Nothing
End Sub

Private Sub txtAgencia_Change()

  If Len(Trim(txtAgencia.Text)) = txtAgencia.MaxLength Then
    If Not bFormating Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub txtAgencia_GotFocus()

  txtAgencia.SelStart = 0
  txtAgencia.SelLength = txtAgencia.MaxLength
End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtAgencia_LostFocus()

  bFormating = True
  txtAgencia.Text = Format(txtAgencia.Text, String(txtAgencia.MaxLength, "0"))
  bFormating = False
End Sub
Private Sub txtBanco_Change()

  If Len(Trim(txtBanco.Text)) = txtBanco.MaxLength Then
    If Not bFormating Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub txtBanco_GotFocus()

  txtBanco.SelStart = 0
  txtBanco.SelLength = txtBanco.MaxLength
End Sub
Private Sub txtBanco_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtBanco_LostFocus()

  bFormating = True
  txtBanco.Text = Format(txtBanco.Text, String(txtBanco.MaxLength, "0"))
  bFormating = False
End Sub
Private Sub txtCheque_Change()

  If Len(Trim(txtCheque.Text)) = txtCheque.MaxLength Then
    If Not bFormating Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub txtCheque_GotFocus()

  txtCheque.SelStart = 0
  txtCheque.SelLength = txtCheque.MaxLength
End Sub
Private Sub txtCheque_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtCheque_LostFocus()

  bFormating = True
  txtCheque.Text = Format(txtCheque.Text, String(txtCheque.MaxLength, "0"))
  bFormating = False
End Sub
Private Sub txtCMC71_Change()

  If Len(Trim(txtCMC71.Text)) = txtCMC71.MaxLength Then
    If Not bFormating Then
      SendKeys "{TAB}"
      DoEvents
    End If
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
Private Sub txtCMC71_LostFocus()

  'Verificar se o campo CMC7_1 é válido
  
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
Private Sub txtComp_Change()

  If Len(Trim(txtComp.Text)) = txtComp.MaxLength Then
    If Not bFormating Then
      SendKeys "{TAB}"
    End If
  End If
End Sub
Private Sub txtComp_GotFocus()

  txtComp.SelStart = 0
  txtComp.SelLength = txtComp.MaxLength
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtComp_LostFocus()

  bFormating = True
  txtComp.Text = Format(txtComp.Text, String(txtComp.MaxLength, "0"))
  bFormating = False
End Sub
Private Sub txtConta_Change()

  If Len(Trim(txtConta.Text)) = txtConta.MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtConta_GotFocus()

  txtConta.SelStart = 0
  txtConta.SelLength = txtConta.MaxLength
End Sub


Private Sub txtConta_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtDV1_Change()

  If Len(Trim(txtDV1.Text)) = txtDV1.MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDV1_GotFocus()

  txtDV1.SelStart = 0
  txtDV1.SelLength = txtDV1.MaxLength
End Sub
Private Sub txtDV1_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtDV2_Change()

  If Len(Trim(txtDV2.Text)) = txtDV2.MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDV2_GotFocus()

  txtDV2.SelStart = 0
  txtDV2.SelLength = txtDV2.MaxLength
End Sub
Private Sub txtDV2_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtDV3_Change()

  If Len(Trim(txtDV3.Text)) = txtDV3.MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDV3_GotFocus()

  txtDV3.SelStart = 0
  txtDV3.SelLength = txtDV3.MaxLength
End Sub
Private Sub txtDV3_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtTipo_Change()

  If Len(Trim(txtTipo.Text)) = txtTipo.MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtTipo_GotFocus()

  txtTipo.SelStart = 0
  txtTipo.SelLength = txtTipo.MaxLength
End Sub
Private Sub txtTipo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Public Sub CalculaDVlinha1()

  On Error GoTo ERRO_CALCULARDVLINHA1

  Dim soma1 As Integer, digito1 As Integer
  Dim soma2 As Integer, digito2 As Integer
  Dim soma3 As Integer, digito3 As Integer
  Dim resto As Integer
  Dim p As Integer
  Dim Campo1, Campo2, Campo3 As String
  Dim unico As Integer, troca As Integer
  Dim dec As Integer

  troca = 0
  unico = 0
  soma1 = 0
  digito1 = 0       'calculado
  soma2 = 0
  digito2 = 0       'calculado
  soma3 = 0
  digito3 = 0       'calculado
  resto = 0

  Campo1 = txtBanco.Text & txtAgencia.Text
  Campo2 = txtComp.Text & txtCheque.Text & txtTipo.Text
  Campo3 = txtConta.Text

  '***************************
  '*** CÁLCULO DO CAMPO 01 ***
  '***************************
  p = Len(Campo1)         'fim do campo 01
  Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
          unico = Mid(Campo1, p, 1) * 2    'multiplica por 2
          troca = 1
      Else
          unico = Mid(Campo1, p, 1) * 1    'multiplica por 1
          troca = 0
      End If

      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
          soma1 = soma1 + unico
      Else
          soma1 = soma1 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
      
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p - 1
      If (p = 0) Then
          Exit Do
      End If
  Loop
 
  '*** Calculo do número decimal para subtração do valor encontrado ***
  If (soma1 > 9) Then
     If Mid(soma1, 2, 1) = 0 Then
        dec = soma1
     Else
        dec = Mid(soma1, 1, 1) + 1 & "0"
     End If
  Else
     If soma1 = 0 Then
        dec = 0
     Else
        dec = 10
     End If
  End If
  '*** Digito verificador calculado pelo módulo 10 ***
  digito1 = dec - soma1
  troca = 0
 
  '***************************
  '*** CÁLCULO DO CAMPO 02 ***
  '***************************
  p = Len(Campo2)         'fim do campo 02
  Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
         unico = Mid(Campo2, p, 1) * 2    'multiplica por 2
         troca = 1
      Else
         unico = Mid(Campo2, p, 1) * 1    'multiplica por 1
         troca = 0
      End If
  
      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
          soma2 = soma2 + unico
      Else
          soma2 = soma2 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
     
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p - 1
      If (p = 0) Then
          Exit Do
      End If
  Loop
  
  '*** Calculo do DV2 do CMC7 ***
  '*** Calculo do número decimal para subtração do valor encontrado ***
  If (soma2 > 9) Then
     If Mid(soma2, 2, 1) = 0 Then
        dec = soma2
     Else
        dec = Mid(soma2, 1, 1) + 1 & "0"
     End If
  Else
     If soma2 = 0 Then
        dec = 0
     Else
        dec = 10
     End If
  End If
  
  '*** Digito verificador calculado pelo módulo 10 ***
  digito2 = dec - soma2
  
  
  '***************************
  '*** CÁLCULO DO CAMPO 03 ***
  '***************************
  p = Len(Campo3)         'fim do campo 03
  Do
      '*** Base 2 (multiplicação dos caracteres por 1 e 2) ***
      If (troca = 0) Then
          unico = Mid(Campo3, p, 1) * 2    'multiplica por 2
          troca = 1
      Else
          unico = Mid(Campo3, p, 1) * 1    'multiplica por 1
          troca = 0
      End If
  
      '*** Verifica se um destes caracteres que foram multiplicados é >= 10.
      '*** Se for maior ou igual, soma o primeiro com o segundo caracter.
      '*** Ex.: 14 = (1 + 4),  18 = (1 + 8), etc...
      If (unico < 10) Then
          soma3 = soma3 + unico
      Else
          soma3 = soma3 + Mid(unico, 1, 1) + Mid(unico, 2, 1)
      End If
     
      '*** aponta para próximo caracter a ser multiplicado ***
      p = p - 1
      If (p = 0) Then
          Exit Do
      End If
  Loop
  
  '*** Calculo do DV3 do CMC7 ***
  '*** Calculo do número decimal para subtração do valor encontrado ***
  If (soma3 > 9) Then
     If Mid(soma3, 2, 1) = 0 Then
        dec = soma3
     Else
        dec = Mid(soma3, 1, 1) + 1 & "0"
     End If
  Else
     If soma3 = 0 Then
        dec = 0
     Else
        dec = 10
     End If
  End If
  
  '*** Digito verificador calculado pelo módulo 10 ***
  digito3 = dec - soma3
  
  txtCMC71.Text = txtBanco & txtAgencia & digito2
  txtCMC72.Text = txtComp & txtCheque & txtTipo
  txtCMC73.Text = digito1 & txtConta & digito3

  Exit Sub

ERRO_CALCULARDVLINHA1:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Linha1.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Private Sub txtValor_GotFocus()

  TxtValor.SelStart = 0
  TxtValor.SelLength = Len(TxtValor.Text)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
