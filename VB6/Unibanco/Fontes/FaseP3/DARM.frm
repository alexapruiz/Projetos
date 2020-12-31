VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form DARM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de DARM"
   ClientHeight    =   1776
   ClientLeft      =   1284
   ClientTop       =   1380
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1776
   ScaleWidth      =   7320
   Begin VB.TextBox TxtIncidencia 
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
      Left            =   2088
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1296
      Width           =   828
   End
   Begin VB.TextBox TxtCCM 
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
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1296
      Width           =   1056
   End
   Begin VB.TextBox TxtTributo 
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
      Left            =   3768
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1296
      Width           =   708
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   5568
      Picture         =   "DARM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   1440
      Picture         =   "DARM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   2268
      Picture         =   "DARM.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   3096
      Picture         =   "DARM.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   3924
      Picture         =   "DARM.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   4752
      Picture         =   "DARM.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Left            =   6396
      Picture         =   "DARM.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   132
      Width           =   816
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
      Height          =   372
      Left            =   5328
      TabIndex        =   3
      Top             =   1296
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tributo"
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
      Left            =   3768
      TabIndex        =   15
      Top             =   1032
      Width           =   624
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Incidência"
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
      Left            =   2040
      TabIndex        =   14
      Top             =   1032
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   168
      Picture         =   "DARM.frx":1546
      Top             =   252
      Width           =   384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "DARM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   684
      TabIndex        =   13
      Top             =   396
      Width           =   480
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
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
      Left            =   5328
      TabIndex        =   12
      Top             =   1032
      Width           =   468
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "CCM"
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
      Left            =   216
      TabIndex        =   11
      Top             =   1032
      Width           =   456
   End
End
Attribute VB_Name = "DARM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryGetDarm As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryAtualizaDARM As rdoQuery
Private qryGetAgenf As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Private Function ValidaCCM() As Boolean

  On Error GoTo ValidaCCM_Err

  ValidaCCM = False

  If (Val(TxtCCM.Text) >= 60000015 And Val(TxtCCM.Text) <= 62500007) Then
    If Verifica_CCM(TxtCCM.Text) = False Then
      MsgBox "C.C.M. incorreto.", vbInformation, App.Title
      TxtCCM.SetFocus
      Exit Function
    End If
  ElseIf (Val(TxtCCM.Text) = 65555554) Or (Val(TxtCCM.Text) = 66666660) Or _
         (Val(TxtCCM.Text) = 77777778) Or (Val(TxtCCM.Text) = 99999998) Or _
         (Val(TxtCCM.Text) = 99999990) Or (Val(TxtCCM.Text) = 99999997) Then

    ValidaCCM = True
    Exit Function
  Else
    If Mid(TxtCCM.Text, 1, 1) = 0 Or Mid(TxtCCM.Text, 1, 1) = 4 Or _
       Mid(TxtCCM.Text, 1, 1) = 5 Or Mid(TxtCCM.Text, 1, 1) = 6 Or _
      Mid(TxtCCM.Text, 1, 1) = 7 Then
      MsgBox "CCM incorreto.", vbInformation, App.Title
      TxtCCM.SetFocus
      Exit Function
    Else
      If Verifica_CCM(TxtCCM.Text) = False Then
        MsgBox "C.C.M. incorreto.", vbInformation, App.Title
        TxtCCM.SetFocus
        Exit Function
      End If
    End If
  End If

  ValidaCCM = True

  Exit Function

ValidaCCM_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar o Código de CCM.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function ValidaIncidencia(ByVal Incidencia As String) As Boolean

   On Error GoTo ERRO_VALIDAINCIDENCIA

   Dim Ano_Inc As Integer
   Dim Mes_Inc As Integer
   Dim Ano_Mov As Integer
   Dim Mes_Mov As Integer
   Dim Dia_Mov As Integer
   Dim ExibeMsg As Boolean

   ValidaIncidencia = False
   ExibeMsg = False

   'Definindo o Ano da Incidencia
   Ano_Inc = Val(Mid(Incidencia, 3, 4))

   'Definindo o Mes da Incidencia
   Mes_Inc = Val(Mid(Incidencia, 1, 2))

   'Definindo o Ano do Movimento
   Ano_Mov = Val(Mid(Geral.DataProcessamento, 1, 4))

   'Definindo o Mes do Movimento
   Mes_Mov = Val(Mid(Geral.DataProcessamento, 5, 2))

   'Definindo o Dia do Movimento
   Dia_Mov = Val(Mid(Geral.DataProcessamento, 7, 2))

   'Verificar se o Ano de Incidencia é igual ao ano do movimento
   If Ano_Inc = Ano_Mov Then
      If Mes_Inc = (Mes_Mov - 1) Then
         'Dentro do prazo
         If Dia_Mov > 10 Then
            'Verificar se data do movimento anterior é menor que 10
            If Not VerificaMovimentoAnterior("10" & Format(Mes_Mov, "00") & Format(Ano_Mov, "0000")) Then
               ExibeMsg = True
            End If
         End If
      ElseIf Mes_Inc < (Mes_Mov - 1) Then
         'Docto Vencido
         ExibeMsg = True
      End If
   ElseIf Ano_Inc = (Ano_Mov - 1) Then
      'Ano de Movimento é maior que ano de Incidencia , verificar mes
      If Mes_Inc = 12 And Mes_Mov = 1 Then
         'Dentro do Prazo
         If Dia_Mov > 10 Then
            'Verificar se data do movimento anterior é menor que 10
            If Not VerificaMovimentoAnterior("10" & Format(Mes_Mov, "00") & Format(Ano_Mov, "0000")) Then
               ExibeMsg = True
            End If
         End If
      Else
         'Documento Vencido
         ExibeMsg = True
      End If
   ElseIf Ano_Inc < (Ano_Mov - 1) Then
      'Documento Vencido
      ExibeMsg = True
   End If

   If ExibeMsg Then
      MsgBox "Este documento só pode ser pago até o dia 10 do mês seguinte ao de Incidência.", vbInformation + vbOKOnly, App.Title
      TxtIncidencia.SetFocus
      TxtIncidencia.SelStart = 0
      TxtIncidencia.SelLength = Len(TxtIncidencia.Text)
      Exit Function
   End If

   ValidaIncidencia = True

   Exit Function

ERRO_VALIDAINCIDENCIA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Incidência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Private Function VerificaMesAno(ByVal Mes As Integer, ByVal Ano As Integer) As Boolean

   On Error GoTo VERIFICAMESANO_ERR

   VerificaMesAno = False

   If Mes > 0 And Mes < 13 Then
      'Mes Válido -> Formatar e Verificar Ano
      Mes = Format(Mes, "00")
      If Ano < 1900 Or Ano > 2100 Then
         'Ano inválido -> Emitir Mensagem
         MsgBox "Ano de Incidência Inválido.", vbInformation, App.Title
         TxtIncidencia.SetFocus
         Exit Function
      Else
         'Ano Válido -> Formatar e preencher campo formatado
         TxtIncidencia.Text = Format(Mes, "00") & "/" & Format(Ano, "0000")
      End If
   Else
      'Mes Inválido -> Emitir Mensagem
      MsgBox "Mês de Incidência Inválido.", vbInformation, App.Title
      TxtIncidencia.SetFocus
      Exit Function
   End If

   VerificaMesAno = True

   Exit Function

VERIFICAMESANO_ERR:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Validar Incidência.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Function
Private Function VerificaMovimentoAnterior(ByVal Data As String) As Boolean

Dim RetAgencia As Integer

   'Validar Agencia
    If Geral.Capa.agefsestado = "" Then
        RetAgencia = ValidaAgencia(Geral.Documento.Agencia, Data, True)
    Else
        If Geral.Capa.agefsstmovi = 9 Then
              RetAgencia = 2
        ElseIf Geral.Capa.agefsstmovi = 0 Then
              RetAgencia = 3
        ElseIf Geral.Capa.agefsstmovi = 2 Then
            'Agencia Aberta -> Verificar data do Movimento Anterior
            If DataAAAAMMDD(Data) <= TransformaDataAAAAMMDD(Geral.Capa.agefsdtmvan) Then
                'A Data de Vencimento é menor ou igual à data do Movimento Anterior -> Não Aceitar
                RetAgencia = 1
            Else
                RetAgencia = 0
            End If
        Else
            RetAgencia = 0
        End If
    End If
   
   'Verificar Retorno da Função
   If RetAgencia = 0 Then
      VerificaMovimentoAnterior = True
   Else
      VerificaMovimentoAnterior = False
   End If
   
End Function
Function VerificaAgenciaSP() As Boolean

  On Error GoTo ERRO_VERIFICAAGENCIASP

  Dim sSql As String
  Dim RsAgenf As rdoResultset

  VerificaAgenciaSP = False

  sSql = Geral.AgenciaCentral

  Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (" & sSql & ")}")

  Set RsAgenf = qryGetAgenf.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  If Not RsAgenf.EOF Then
    'Encontrou a Agencia -> Verificar se é de SP
    If RsAgenf!agefsestado <> "SP" Then
      MsgBox "A Agência de Origem não aceita DARM.", vbInformation, App.Title
      Exit Function
    End If
  Else
    MsgBox "A Agência de Origem não está cadastrada.", vbInformation, App.Title
    Exit Function
  End If

  VerificaAgenciaSP = True

  Exit Function

ERRO_VERIFICAAGENCIASP:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Agência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdConfirmar_Click()

  If SalvaDARM Then
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

  Call PesquisaDARM
End Sub

Public Function VerificaTributo(pvsCodigoTributo) As Boolean
   
  '--------- MODULO 11 (2 BASE 9) --------------
  ' Esta rotina serve para conferir o codigo da txt_tributoeita

  Dim soma As Integer, resto As Integer
  Dim digito_11 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim bOk As Boolean

  bOk = True

  soma = 0
  resto = 0
  digito_11 = 0        'calculado pelo módulo 11
  digito_rv = ""       'caracter digitado pelo operador

  soma = soma + Mid(pvsCodigoTributo, 1, 1) * 5
  soma = soma + Mid(pvsCodigoTributo, 2, 1) * 4
  soma = soma + Mid(pvsCodigoTributo, 3, 1) * 3
  soma = soma + Mid(pvsCodigoTributo, 4, 1) * 2

  resto = soma Mod 11        'resto da divisão
  digito_11 = 11 - resto     'digito verificador

  '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
  If (digito_11 = 11) Or (digito_11 = 10) Then
    digito_11 = 0
  End If

  digito_rv = Mid(pvsCodigoTributo, 5, 1)    'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
    bOk = False           'digito não confere
  End If

  VerificaTributo = bOk
End Function
Public Function Verifica_CCM(ccm) As Boolean

  Dim soma As Integer, resto As Integer
  Dim digito_11 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim bOk As Boolean

  bOk = True

  soma = 0
  resto = 0
  digito_11 = 0        'calculado pelo módulo 11
  digito_rv = ""       'caracter digitado pelo operador

  'número do CCM: (7+1)                 0 0 0 0 0 0 0 - D
  '                                     x x x x x x x x x
  'multiplica da direita para esquerda: 8 7 6 5 4 3 2

  peso = 2    'começa multiplicar da direita para esquerda
  p = 7      'tamanho do campo sem o digito

  ccm = Format(ccm, "00000000")

  Do
    '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
    soma = soma + Mid(ccm, p, 1) * peso
    p = p - 1            'ponteiro
    peso = peso + 1      'peso
    If (peso = 9) Then
      peso = 2
    End If
    If (p = 0) Then
      Exit Do
    End If
  Loop

  resto = soma Mod 11      'resto da divisão
  digito_11 = 11 - resto      'digito verificador

  '*** se o calculo for igual a 0 ou 1, muda-se para 0 ***
  If digito_11 = 10 Or digito_11 = 11 Then
    digito_11 = 0
  End If

  digito_rv = Mid(ccm, 8, 1)  'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
    bOk = False               'digito não confere
    Verifica_CCM = bOk
    Exit Function
  End If

  Verifica_CCM = bOk

End Function
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub

Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub



Function SalvaDARM() As Boolean

On Error GoTo ERRO_SALVADARM
  
  Dim strEncripta   As String
  
  SalvaDARM = False

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Formatando o Campo LEITURA
    Geral.Documento.Leitura = "816" & Format(Val(Desformata_Valor(TxtValor.Text)), "000000000000") & _
            "000006" & Format(TxtCCM.Text, "00000000") & Format(TxtTributo.Text, "00000") & _
            Format(Mid(TxtIncidencia.Text, 1, 2) & Mid(TxtIncidencia.Text, 6, 2), "0000") & "112300"

    'Formatando o Campo 'CÓDIGO TRIBUTO'
    TxtTributo.Text = Format(TxtTributo.Text, "00000")

    'Validar CCM
    If Not ValidaCCM Then Exit Function

    'Validar Incidencia
    If Not ValidaIncidencia(Mid(TxtIncidencia.Text, 1, 2) & Mid(TxtIncidencia.Text, 4)) Then
      TxtIncidencia.SetFocus
      Exit Function
    End If

    'Validar Código de Tributo
    If Not VerificaTributo(TxtTributo.Text) Then
      MsgBox "Código de Tributo Inválido.", vbInformation, App.Title
      TxtTributo.SetFocus
      Exit Function
    End If

    'Verificar se a Agencia é de SP
    If Not VerificaAgenciaSP Then Exit Function

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 15 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(15, CStr(Val(TxtCCM.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVADARM
    
    'Atualizar / Inserir DARF Preto
    With qryAtualizaDARM
      .rdoParameters(0) = Geral.DataProcessamento              'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto              'IdDocto
      .rdoParameters(2) = Mid(TxtIncidencia.Text, 4, 4) & _
                          Mid(TxtIncidencia.Text, 1, 2)        'Incidencia
      .rdoParameters(3) = TxtCCM.Text                          'CCM
      .rdoParameters(4) = TxtTributo.Text                      'Tributo
      .rdoParameters(5) = Val(TxtValor.Text) / 100             'Valor
      .rdoParameters(6) = Geral.Documento.Leitura              'Leitura
      .rdoParameters(7) = 15                                   'TipoDocto
      .rdoParameters(8) = strEncripta                          'Autenticacao digital
      .Execute
    End With

    SalvaDARM = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.TipoDocto = 15
  End If

  Exit Function

ERRO_SALVADARM:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do DARM.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PesquisaDARM()

  On Error GoTo ERRO_PESQUISADARM

  Dim sSql As String
  Dim RsDARM As rdoResultset

  'Preencher os campos do DARF , caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetDarm = Geral.Banco.CreateQuery("", "{call GetDARM (" & sSql & ")}")

  Set RsDARM = qryGetDarm.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsDARM.EOF Then
    'Encontrou o DARM -> Preencher os campos
    TxtCCM.Text = RsDARM!ccm
    TxtIncidencia.Text = Mid(RsDARM!Incidencia, 5, 2) & Mid(RsDARM!Incidencia, 1, 4)
    TxtTributo.Text = Format(RsDARM!Tributo, "00000")
    TxtValor.Text = Val(RsDARM!Valor * 100)

    'Posicionar o Foco no campo 'VALOR'
    TxtValor.SetFocus
  Else
    'Posicionar o Foco no campo 'PERIODO APURAÇÃO'
    TxtCCM.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PESQUISADARM:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados do DARM.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CamposOK() As Boolean

  On Error GoTo ERRO_CAMPOSOK

  CamposOK = False

   'CCM
   If Len(Trim(TxtCCM.Text)) = 0 Then
      MsgBox "Informe o CCM do DARM.", vbInformation, App.Title
      TxtCCM.SetFocus
      Exit Function
   End If

   'Verificar Incidencia zerada
   If Len(Trim(TxtIncidencia.Text)) = 0 Then
      MsgBox "Informe a Incidência.", vbInformation, App.Title
      TxtIncidencia.SetFocus
      Exit Function
   End If

   If Not ValidaData Then
      'MsgBox "Informe a Incidência.", vbInformation, App.Title
      TxtIncidencia.SetFocus
      Exit Function
   End If

  'Tributo
   If Len(Trim(TxtTributo.Text)) = 0 Then
      MsgBox "Informe o Código do Tributo.", vbInformation, App.Title
      TxtTributo.SetFocus
      Exit Function
   End If

  'Valor
  If Val(TxtValor.Text) = 0 Then
    MsgBox "Informe o Valor do Documento.", vbInformation, App.Title
    TxtValor.SetFocus
    Exit Function
  End If

  CamposOK = True

  Exit Function

ERRO_CAMPOSOK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar os valores dos campos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub AjustesIniciais()

  'Setando as Variáveis do RDO
  Set qryGetDarm = Geral.Banco.CreateQuery("", "{? = call GetDARM (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryAtualizaDARM = Geral.Banco.CreateQuery("", "{call AtualizaDARM (?,?,?,?,?,?,?,?,?)}")
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
  Set qryGetDarm = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryAtualizaDARM = Nothing
End Sub
Private Sub TxtCCM_GotFocus()
  TxtCCM.SelStart = 0
  TxtCCM.SelLength = Len(TxtCCM.Text)
End Sub
Private Sub TxtCCM_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtCCM_LostFocus()
'* Valida CCM *'
    If Len(Trim(TxtCCM)) = 0 Then Exit Sub
        If Not IsNumeric(TxtCCM) Then
            MsgBox "Código do CCM inválido, Redigite.", vbInformation, App.Title
            TxtCCM.Text = ""
            TxtCCM.SetFocus
        End If
End Sub
Private Sub TxtIncidencia_GotFocus()
  TxtIncidencia.SelStart = 0
  TxtIncidencia.SelLength = Len(TxtIncidencia.Text)
End Sub
Private Sub TxtIncidencia_KeyPress(KeyAscii As Integer)

  If InStr(TxtIncidencia.Text, "/") = 0 And Len(Trim(TxtIncidencia.Text)) = 6 Then
    'Verificar se a tecla atual é um numero
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 And TxtIncidencia.SelLength = 0 Then
      KeyAscii = 0
    End If
  End If

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 47 Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtIncidencia_LostFocus()

   'Verificar se o usuário informou Mes e Ano de Incidencia
   On Error GoTo ERRO_INCIDENCIALOSTFOCUS

   If Not ValidaData Then Exit Sub

   Exit Sub

ERRO_INCIDENCIALOSTFOCUS:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Validar Incidência.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Sub
Private Function ValidaData() As Boolean

   On Error GoTo ValidaData_Err

   Dim Data As String
   Dim Pos As Integer

   Dim Mes As String
   Dim Ano As String

   ValidaData = False

   If Len(Trim(TxtIncidencia.Text)) = 0 Then Exit Function

   With TxtIncidencia
      .Text = Trim(.Text)
      If .Text <> "" Then
         Pos = InStr(1, .Text, "/", vbTextCompare)
         If Pos > 0 Then
            'O Usuário digitou barra
            Mes = Mid(.Text, 1, Pos - 1)
            Ano = Mid(.Text, Pos + 1)
            If Not VerificaMesAno(Mes, Ano) Then Exit Function

            Data = Mid(.Text, 1, Pos - 1) & Mid(.Text, Pos + 1, Len(.Text) - Pos)
         Else
            'O Usuário não informou Data
            If Len(.Text) <> 6 Then
               MsgBox "O Campo Incidência deve estar no formato MM/AAAA.", vbInformation, App.Title
               .SelStart = 0
               .SelLength = Len(.Text)
               .SetFocus
               Exit Function
            Else
               Data = .Text
               Mes = Mid(.Text, 1, 2)
               Ano = Mid(.Text, 3)

               If Not VerificaMesAno(Mes, Ano) Then Exit Function
            End If
         End If

         If (Not IsNumeric(Data)) Or Val(Data) < 0 Or Mid(Data, 3) > 2100 Or Mid(Data, 1, 2) > 13 Then
            MsgBox "Incidência Inválida.", vbInformation, App.Title
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
         Else
            .Text = Format(Data, "00/0000")
         End If
      End If
  End With
  
  ValidaData = True
  
  Exit Function

ValidaData_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Incidência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub TxtTributo_GotFocus()
  TxtTributo.SelStart = 0
  TxtTributo.SelLength = Len(TxtTributo.Text)
End Sub
Private Sub TxtTributo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtTributo_LostFocus()
'* Valida Código de Tributo do DARM *'
    
    If Len(Trim(TxtTributo)) = 0 Then Exit Sub
        If Not IsNumeric(TxtTributo) Then
            MsgBox "Código do Tributo inválido, Redigite", vbInformation, App.Title
            TxtTributo.Text = ""
            TxtTributo.SetFocus
            Exit Sub
        End If
        
End Sub

Private Sub TxtTributo_Validate(Cancel As Boolean)

   If Len(Trim(TxtTributo.Text)) > 0 And Val(TxtTributo.Text) <> 0 Then
      'Validar Código de Tributo
      TxtTributo.Text = Format(TxtTributo.Text, "00000")
      If Not VerificaTributo(TxtTributo.Text) Then
         MsgBox "Código de Tributo Inválido.", vbInformation, App.Title
         Cancel = True
         TxtTributo.SelStart = 0
         TxtTributo.SelLength = Len(TxtTributo.Text)
      End If
   End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
