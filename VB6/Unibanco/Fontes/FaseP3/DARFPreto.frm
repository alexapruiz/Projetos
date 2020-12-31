VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form DARFPreto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de DARF Preto"
   ClientHeight    =   2436
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   8136
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2436
   ScaleWidth      =   8136
   Begin VB.PictureBox PicRef 
      Height          =   816
      Left            =   72
      ScaleHeight     =   768
      ScaleWidth      =   7908
      TabIndex        =   26
      Top             =   2508
      Visible         =   0   'False
      Width           =   7956
      Begin VB.CommandButton CmdCancelarRef 
         Caption         =   "&Cancelar"
         Height          =   324
         Left            =   6516
         TabIndex        =   29
         Top             =   240
         Width           =   1296
      End
      Begin VB.Frame Frame1 
         Caption         =   "Confirmação do Campo 'Referência'"
         Height          =   696
         Left            =   36
         TabIndex        =   27
         Top             =   12
         Width           =   6384
         Begin VB.TextBox TxtReferencia2 
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
            Left            =   3492
            MaxLength       =   17
            TabIndex        =   28
            Top             =   240
            Width           =   1944
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BackStyle       =   0  'Transparent
            Caption         =   "Confirmar Referência"
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
            Left            =   1512
            TabIndex        =   30
            Top             =   288
            Width           =   1908
         End
      End
   End
   Begin VB.TextBox TxtReferencia 
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
      Left            =   4710
      MaxLength       =   17
      TabIndex        =   3
      Top             =   1284
      Width           =   1944
   End
   Begin VB.TextBox TxtReceita 
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
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1284
      Width           =   1152
   End
   Begin DATEEDITLib.DateEdit TxtPeriodo 
      Height          =   372
      Left            =   108
      TabIndex        =   0
      Top             =   1284
      Width           =   1296
      _Version        =   65537
      _ExtentX        =   2286
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
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   6384
      Picture         =   "DARFPreto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   696
      Left            =   2256
      Picture         =   "DARFPreto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   696
      Left            =   3084
      Picture         =   "DARFPreto.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   696
      Left            =   3912
      Picture         =   "DARFPreto.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      Height          =   696
      Left            =   4740
      Picture         =   "DARFPreto.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      Height          =   696
      Left            =   5568
      Picture         =   "DARFPreto.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   696
      Left            =   7212
      Picture         =   "DARFPreto.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   132
      Width           =   816
   End
   Begin VB.TextBox TxtCGCCPF 
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
      Left            =   1482
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1284
      Width           =   1920
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtPrincipal 
      Height          =   372
      Left            =   72
      TabIndex        =   5
      Top             =   1992
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtMulta 
      Height          =   372
      Left            =   2100
      TabIndex        =   6
      Top             =   2004
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtJuros 
      Height          =   372
      Left            =   4128
      TabIndex        =   7
      Top             =   2004
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
   Begin DATEEDITLib.DateEdit TxtVencimento 
      Height          =   372
      Left            =   6732
      TabIndex        =   4
      Top             =   1284
      Width           =   1296
      _Version        =   65537
      _ExtentX        =   2286
      _ExtentY        =   656
      _StockProps     =   93
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
      Height          =   372
      Left            =   6156
      TabIndex        =   8
      Top             =   2004
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
      Locked          =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
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
      Left            =   6768
      TabIndex        =   25
      Top             =   1020
      Width           =   1056
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
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
      Left            =   4800
      TabIndex        =   24
      Top             =   1020
      Width           =   972
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Receita"
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
      Left            =   3516
      TabIndex        =   23
      Top             =   1020
      Width           =   684
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   168
      Picture         =   "DARFPreto.frx":1546
      Top             =   252
      Width           =   384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "DARF Preto"
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
      Left            =   684
      TabIndex        =   22
      Top             =   396
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
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
      Left            =   6228
      TabIndex        =   21
      Top             =   1752
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Juros "
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
      Left            =   4176
      TabIndex        =   20
      Top             =   1752
      Width           =   552
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Multa"
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
      Left            =   2160
      TabIndex        =   19
      Top             =   1752
      Width           =   492
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Principal"
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
      TabIndex        =   18
      Top             =   1752
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "CPF / CGC"
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
      Left            =   1560
      TabIndex        =   17
      Top             =   1020
      Width           =   996
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Apuração"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1020
      Width           =   852
   End
End
Attribute VB_Name = "DARFPreto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryGetDARFPreto As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryAtualizaDARFPreto As rdoQuery
Private qryGetValidaDARFPreto As rdoQuery

'Declaração de Variáveis de trabalho
Private ValidaDARF As TpValidaDARF
Private mForm As Form
Public Alterou As Boolean
Public AlteraValor As Boolean

Private Type TpValidaDARF
  sTipoDocumento As String * 1
  sReferencia As String * 1
  sDataInicial As String * 8
  sDataFinal As String * 8
  nValorLimite  As Currency
  AnoLimiteInicial As Integer
  AnoLimiteFinal As Integer
  ExcessaoDocto As String * 1
  DataLimiteApur As Long
End Type

'Validação de informações do DARF SIMPLES
Private RegraValidaDARF() As tpValidacao
Private Type tpValidacao
    CodigoReceitaNaoAceito  As Long
    DataBase1               As Long
    DataBase2               As Long
    'ValorMinimoDocumento    As Currency
End Type
'

Sub CalculaValorFinal()

  On Error GoTo ERRO_CALCULAVALORFINAL

  Dim Valor As Currency

  '----------------------------   Calcular o Valor Final do Documento --------------------------

  'Verificar se foi informado o Valor Principal
  If Val(TxtPrincipal.Text) = 0 Then
    Valor = 0
  Else
    Valor = TxtPrincipal.Text
  End If

  'Verificar se foi informado Multa
  If Val(TxtMulta.Text) <> 0 Then
    Valor = Valor + Val(TxtMulta.Text)
  End If

  'Verificar se foi informado Juros
  If Val(TxtJuros.Text) <> 0 Then
    Valor = Valor + Val(TxtJuros.Text)
  End If

  'Transportar o Valor Final para a tela
  TxtValor.Text = Valor

  Exit Sub

ERRO_CALCULAVALORFINAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Final do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Function VerificaExcessao(ByVal CGCCPF As String, ByVal TipoDoc As Integer) As Boolean

   VerificaExcessao = False

   Select Case TipoDoc
      Case 2
         'Verificar se Campo ExcessaoDocto = 1 ou 2
         If ValidaDARF.ExcessaoDocto = "1" And Val(CGCCPF) = 191 Then
            VerificaExcessao = True
            Exit Function
         End If

         If ValidaDARF.ExcessaoDocto = "2" And (TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase1 Or TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase2) Then
            'Verificar se o numero digitado é um CNPJ válido
            CGCCPF = Format(CGCCPF, String(15, "0"))
            If VerificaCGC(CGCCPF) Then
               VerificaExcessao = True
               Exit Function
            End If
         End If

      Case 3
         'Verificar se Campo ExcessaoDocto = 1 ou 3
         If ValidaDARF.ExcessaoDocto = "1" And Val(CGCCPF) = 191 Then
            VerificaExcessao = True
            Exit Function
         End If

         If ValidaDARF.ExcessaoDocto = "3" And (TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase1 Or TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase2) Then
            'Verificar se o numero digitado é um CNPJ válido
            CGCCPF = Format(CGCCPF, String(11, "0"))
            If VerificaCPF(CGCCPF) Then
               VerificaExcessao = True
               Exit Function
            End If
         End If

      Case 5
         'Verificar se Campo ExcessaoDocto = 2 ou 4
         If ValidaDARF.ExcessaoDocto = "2" Then
           'Se Data Apur. = 01-01-80 ou 08-08-80 , aceitar CNPJ Matriz e Filial
           If (TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase1 Or TxtPeriodo.InverseText = RegraValidaDARF(1).DataBase2) Then
             VerificaExcessao = True
           End If

         ElseIf ValidaDARF.ExcessaoDocto = "4" Then
            'Verificar se a data de apuração é maior que a Data Limite
            If TxtPeriodo.InverseText > ValidaDARF.DataLimiteApur Then
               'Só será aceito CNPJ da Matriz
               CGCCPF = Format(CGCCPF, String(15, "0"))
               If (Mid(CGCCPF, 10, 4) = "0001" And VerificaCGC(CGCCPF) = True) Then
                  VerificaExcessao = True
                  Exit Function
               End If
            Else
               'Será aceito CNPJ de Matriz e Filial
               VerificaExcessao = True
               Exit Function
            End If
         Else
            VerificaExcessao = True
         End If

   End Select

End Function
Private Function VerificaReferenciaDV1(ByVal REF As String) As Boolean

  On Error GoTo VerificaReferenciaDV1_Err

  Dim soma As Integer
  Dim resto As Integer
  Dim digito_11 As Integer
  Dim digito_rv As String
  Dim p As Integer
  Dim peso As Integer
  Dim bOk As Boolean

  soma = 0
  resto = 0
  digito_11 = 0                   'Calculado pelo Módulo 11
  digito_rv = ""                  'Digito informado pelo usuário
  peso = 2                        'Peso inicial
  p = 16                          'Primeiro byte da multiplicação

  bOk = True                      ' default - ok

  Do
    ' *********************************************************
    ' * Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) *
    ' *********************************************************
    soma = soma + Mid(REF, p, 1) * peso
    p = p - 1
    peso = peso + 1

    If peso = 10 Then
      peso = 2
    End If

    If p = 0 Then
      Exit Do
    End If

    DoEvents
  Loop

  resto = soma Mod 11             'Resto da divisão
  digito_11 = 11 - resto          'Digito Verificador

  If digito_11 = 11 Or digito_11 = 10 Then
    digito_11 = 0
  End If

  digito_rv = Mid(REF, 17, 1)     ' digito verificador

  If CStr(digito_11) <> (digito_rv) Then
    bOk = False                 ' digito não confere
  End If

  VerificaReferenciaDV1 = bOk

  Exit Function

VerificaReferenciaDV1_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Número de Referência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificaReferenciaDV2(ByVal REF As String) As Boolean

  On Error GoTo VerificaReferenciaDV2_Err

  Dim soma As Integer
  Dim resto As Integer
  Dim Digito_1 As Integer
  Dim Digito_2 As Integer
  Dim digito_rv As String
  Dim p As Integer
  Dim peso As Integer
  Dim bOk As Boolean

  'Calculando primeiro DV
  soma = 0
  resto = 0
  Digito_1 = 0                    'Primeiro dígito verificador
  digito_rv = ""                  'Digito informado pelo usuário
  peso = 2                        'Peso Inicial
  p = 15                          'Primeiro byte da multiplicação

  bOk = True                      ' default - ok

  Do
    ' *********************************************************
    ' * Peso de 2 a N (multiplicação dos caracteres de 2 a N) *
    ' *********************************************************
    soma = soma + Mid(REF, p, 1) * peso
    p = p - 1                   ' ponteiro
    peso = peso + 1             ' peso

    If (p = 0) Then
        Exit Do
    End If

    DoEvents
  Loop

  resto = soma Mod 11             'Resto da divisão
  Digito_1 = 11 - resto           'Digito Verificador

  If Digito_1 = 11 Or Digito_1 = 10 Or Digito_1 = 1 Then
    Digito_1 = 0
  End If

  'Calculando segundo DV
  soma = 0
  resto = 0
  Digito_2 = 0                    'Segundo dígito verificador
  digito_rv = ""                  'Digito informado pelo usuário
  peso = 2                        'Peso inicial para multiplicação
  p = 16                          'Primeiro byte da soma (da direita para a esquerda)

  Do
    ' *********************************************************
    ' * Peso de 2 a N (multiplicação dos caracteres de 2 a N) *
    ' *********************************************************
    soma = soma + Mid(REF, p, 1) * peso
    p = p - 1
    peso = peso + 1

    If (p = 0) Then
        Exit Do
    End If

    DoEvents
  Loop

  resto = soma Mod 11             'Resto da divisão
  Digito_2 = 11 - resto           'Segundo Digito Verificador

  If Digito_2 = 10 Or Digito_2 = 1 Then
    Digito_2 = 0
  End If

  digito_rv = Mid(REF, 16, 2)     'Digito Verificador digitado pelo usuário

  If resto = 0 Then
    'Verificar se bate com DV = 0
    If CStr(Digito_1 & "0") <> (digito_rv) Then
      'Não bate -> verificar com DV = 1
      If CStr(Digito_1 & "1") <> (digito_rv) Then
        bOk = False
      End If
    End If
  Else
    If CStr(Digito_1 & Digito_2) <> (digito_rv) Then
      bOk = False                 ' digito não confere
    End If
  End If

  VerificaReferenciaDV2 = bOk

  Exit Function

VerificaReferenciaDV2_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Número de Referência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Private Sub CmdCancelarRef_Click()

  Me.Height = 2820
  PicRef.Visible = False
  TxtReferencia2.Text = ""
End Sub

Private Sub cmdConfirmar_Click()

  If SalvaDARFPreto Then
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

    If Not CargaRegraValidaDARF Then
        Alterou = False
        Me.Hide
        Exit Sub
    End If

  Call AjustesIniciais

  Call PesquisaDARFPreto
End Sub
Private Function ValidaReceita() As Boolean

   On Error GoTo ERRO_VALIDARECEITA

    Dim sDataHoje As String
    Dim sDataInicio As String
    Dim sDataFim As String
    Dim RsGetValidaDarfPreto As rdoResultset
    Dim nCont As Long
    
   ValidaReceita = False

   TxtReceita.Text = Format(TxtReceita.Text, "0000")

   'Verifica se o Código da Receita é Válido
   If TxtReceita.Text = "6106" Or TxtReceita.Text = "2593" Or TxtReceita.Text = "4570" Then
      MsgBox "Codigo de Receita inválido para este formulário ! Utilize DARF Simples para digitá-lo .", vbInformation, App.Title
      TxtValor.SetFocus
      Exit Function
   End If

   'Verifica o Digito do Código da Receita
   If Not VerificaReceita(TxtReceita.Text) Then
      MsgBox "Código da Receita Inválido.", vbInformation, App.Title
      TxtReceita.SetFocus
      Exit Function
   End If

    'Verifica se o Código da Receita é Válido
    For nCont = 1 To UBound(RegraValidaDARF)
        If TxtReceita.Text = RegraValidaDARF(nCont).CodigoReceitaNaoAceito Then
           MsgBox "Código de Receita exclusivo do Documento de Depósito - DJE." & Chr(13) & Space(6) & "Recolhimento somente na Caixa Econômica Federal.", vbInformation, App.Title
           TxtReceita.SetFocus
            Exit Function
        End If
    Next

   'Verifica se o Código da Receita está Cadastrado
   With qryGetValidaDARFPreto
      .rdoParameters(0) = Val(TxtReceita.Text)
   End With

   Set RsGetValidaDarfPreto = qryGetValidaDARFPreto.OpenResultset(rdOpenStatic, rdConcurReadOnly)

   If RsGetValidaDarfPreto.EOF Then
        'Para código da receita inexistente na tabela de DarfPreto, verificar apenas CPF/CNPJ
        ValidaDARF.sTipoDocumento = 1
   Else
        With ValidaDARF
           .sTipoDocumento = RsGetValidaDarfPreto!TipoDocumento & ""
           .sReferencia = RsGetValidaDarfPreto!Referencia & ""
           .sDataInicial = Format(RsGetValidaDarfPreto!DataInicial, "00000000") & ""
           .sDataFinal = Format(RsGetValidaDarfPreto!DataFinal, "00000000") & ""
           .nValorLimite = RsGetValidaDarfPreto!ValorLimite & ""
           .AnoLimiteInicial = Val(RsGetValidaDarfPreto!AnoLimiteInicial & "")
           .AnoLimiteFinal = Val(RsGetValidaDarfPreto!AnoLimiteFinal & "")
           .ExcessaoDocto = RsGetValidaDarfPreto!ExcessaoDocto & ""
           .DataLimiteApur = Val(RsGetValidaDarfPreto!DataLimiteApur & "")
        End With

        'Verificar se o periodo de apuracao é igual a "01011980" ou "08081980"
        If TxtPeriodo.InverseText <> RegraValidaDARF(1).DataBase1 And TxtPeriodo.InverseText <> RegraValidaDARF(1).DataBase2 Then
           'Verifica se este Código de Receita está sendo recebido
           sDataHoje = Geral.DataProcessamento
           sDataInicio = Mid(ValidaDARF.sDataInicial, 5, 4) & Mid(ValidaDARF.sDataInicial, 3, 2) & Mid(ValidaDARF.sDataInicial, 1, 2)
           sDataFim = Mid(ValidaDARF.sDataFinal, 5, 4) & Mid(ValidaDARF.sDataFinal, 3, 2) & Mid(ValidaDARF.sDataFinal, 1, 2)
        
           If Trim(sDataHoje) < Trim(sDataInicio) Or _
              (sDataFim <> "00000000" And Trim(sDataHoje) > Trim(sDataFim)) Then
              MsgBox "Data inválida para este código de receita.", vbInformation, App.Title
              TxtPeriodo.SetFocus
              Exit Function
           End If
        End If

        'Validação do Valor
        If ValidaDARF.nValorLimite > 0 Then
            If CCur(TxtValor.Text / 100) < ValidaDARF.nValorLimite Then
                MsgBox "Valor Arrecadado não pode ser inferior a R$ " & Trim(FormataValor(ValidaDARF.nValorLimite, 10)) & " .", vbInformation + vbOKOnly, App.Title
                TxtPrincipal.SetFocus
                Exit Function
            End If
        End If

        'Verificar Ano de Apuração
        If (Val(Mid(TxtPeriodo.InverseText, 1, 4)) < ValidaDARF.AnoLimiteInicial _
              And ValidaDARF.AnoLimiteInicial <> 0) Or _
           (Val(Mid(TxtPeriodo.InverseText, 1, 4)) > ValidaDARF.AnoLimiteFinal _
              And ValidaDARF.AnoLimiteFinal <> 0) Then
           MsgBox "Ano de Apuração Inválido.", vbInformation + vbOKOnly, App.Title
           TxtPeriodo.SetFocus
           Exit Function
        End If

        'Verificar Ano de Vencimento
        If (Val(Mid(TxtVencimento.InverseText, 1, 4)) < ValidaDARF.AnoLimiteInicial _
              And ValidaDARF.AnoLimiteInicial <> 0) Or _
           (Val(Mid(TxtVencimento.InverseText, 1, 4)) > ValidaDARF.AnoLimiteFinal _
              And ValidaDARF.AnoLimiteFinal <> 0) Then
           MsgBox "Ano de Vencimento Inválido.", vbInformation + vbOKOnly, App.Title
           TxtVencimento.SetFocus
           Exit Function
        End If
    End If
   
   'Verificar Campo 'TipoDocumento' e 'IndicaExcessaoDocumento'
   Select Case ValidaDARF.sTipoDocumento
      Case 1
         'CNPJ ou CPF com DV
         TxtCGCCPF.Text = Format(TxtCGCCPF.Text, String(15, "0"))
         If Not VerificaCGC(TxtCGCCPF.Text) Then
            'Não bateu DV para CNPJ -> Verificar CPF
            If Len(Trim(Val(TxtCGCCPF.Text))) <= 11 Then
               TxtCGCCPF.Text = Format(TxtCGCCPF.Text, String(11, "0"))
               If Not VerificaCPF(TxtCGCCPF.Text) Then
                  'Não bateu DV para CPF
                  MsgBox "Código de CNPJ / CPF inválido.", vbInformation + vbOKOnly, App.Title
                  TxtCGCCPF.SetFocus
                  Exit Function
               End If
            Else
               'CPF / CNPJ Inválido
               MsgBox "Código de CNPJ / CPF inválido.", vbInformation + vbOKOnly, App.Title
               TxtCGCCPF.SetFocus
               Exit Function
            End If
         End If

      Case 2
         'CPF com DV
         If Len(Trim(Val(TxtCGCCPF.Text))) > 11 Then
            'Verificar Campo ExcessaoDocto
            If Not VerificaExcessao(TxtCGCCPF.Text, 2) Then
               MsgBox "CPF / CNPJ Inválido.", vbInformation + vbOKOnly, App.Title
               TxtCGCCPF.SetFocus
               Exit Function
            End If
         Else
            TxtCGCCPF.Text = Format(TxtCGCCPF.Text, String(11, "0"))
            If Not VerificaCPF(TxtCGCCPF.Text) Then
               MsgBox "CPF / CNPJ Inválido.", vbInformation + vbOKOnly, App.Title
               TxtCGCCPF.SetFocus
               Exit Function
            End If
         End If
      Case 3
         'CNPJ com DV
         TxtCGCCPF.Text = Format(TxtCGCCPF.Text, String(15, "0"))
         If (Not VerificaCGC(TxtCGCCPF.Text)) Or (Right(TxtCGCCPF.Text, 3) = "191") Then
            'Verificar Campo ExcessaoDocto
            If Not VerificaExcessao(TxtCGCCPF.Text, 3) Then
               MsgBox "CPF / CNPJ Inválido.", vbInformation + vbOKOnly, App.Title
               TxtCGCCPF.SetFocus
               Exit Function
            End If
         End If
      Case 5
         'Somente CNPJ da Matriz (com DV)
         TxtCGCCPF.Text = Format(TxtCGCCPF.Text, String(15, "0"))

         'Validar DV
         If VerificaCGC(TxtCGCCPF.Text) Then
            If Mid(TxtCGCCPF.Text, 10, 4) <> "0001" Then
               If Not VerificaExcessao(TxtCGCCPF.Text, 5) Then
                  MsgBox "CNPJ Inválido.", vbInformation + vbOKOnly, App.Title
                  TxtCGCCPF.SetFocus
                  Exit Function
               End If
            End If
         Else
            MsgBox "CNPJ Inválido.", vbInformation + vbOKOnly, App.Title
            TxtCGCCPF.SetFocus
            Exit Function
         End If
   End Select

  'Verifica CPF 111.111.111-11 a 999.999.999-99
  If TxtCGCCPF.Text = "11111111111" Or _
    TxtCGCCPF.Text = "22222222222" Or _
    TxtCGCCPF.Text = "33333333333" Or _
    TxtCGCCPF.Text = "44444444444" Or _
    TxtCGCCPF.Text = "55555555555" Or _
    TxtCGCCPF.Text = "66666666666" Or _
    TxtCGCCPF.Text = "77777777777" Or _
    TxtCGCCPF.Text = "88888888888" Or _
    TxtCGCCPF.Text = "99999999999" Or _
    TxtCGCCPF.Text = "00000000000" Then
    MsgBox "Código de CPF não permitido.", vbInformation, App.Title
    TxtCGCCPF.SetFocus
    Exit Function
  End If

  'Verifica CNPJ 00.000.000 a 99.999.999
  If Mid(TxtCGCCPF.Text, 2, 8) = "00000000" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "11111111" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "22222222" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "33333333" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "44444444" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "55555555" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "66666666" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "77777777" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "88888888" Or _
     Mid(TxtCGCCPF.Text, 2, 8) = "99999999" Then
    If Len(Trim(TxtCGCCPF.Text)) = 15 And Right(TxtCGCCPF.Text, 3) <> "191" Then
      MsgBox "Código de CNPJ não permitido.", vbInformation, App.Title
      TxtCGCCPF.SetFocus
      Exit Function
    End If
  End If

  'Verificar se numero de ordem igual A 0000
  If Mid(TxtCGCCPF.Text, 10, 4) = "0000" Then
    MsgBox "Numero de ordem não permitido.", vbInformation, App.Title
    TxtCGCCPF.SetFocus
    Exit Function
  End If

  ValidaReceita = True

  Exit Function

ERRO_VALIDARECEITA:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Validar Código da Receita.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Function
Public Function VerificaReceita(ByVal REC As String) As Boolean

    Dim soma As Integer
    Dim resto As Integer
    Dim digito_11 As Integer
    Dim digito_rv As String
    Dim bOk As Boolean

    bOk = True                      ' digito confere - default
    
    soma = 0
    resto = 0
    digito_11 = 0                   ' calculado pelo módulo 11
    digito_rv = ""                  ' caracter digitado pelo operador
    
    soma = soma + Mid(REC, 1, 1) * 8
    soma = soma + Mid(REC, 2, 1) * 4
    soma = soma + Mid(REC, 3, 1) * 2
    
    resto = soma Mod 11             ' resto da divisão
    digito_11 = 11 - resto          ' digito verificador
   
    ' *****************************************************
    ' * Se o Cálculo for Igual a 10 ou 11, muda-se para 0 *
    ' *****************************************************
    If digito_11 = 11 Or digito_11 = 10 Then
        digito_11 = 0
    End If

    digito_rv = Mid(REC, 4, 1)      ' digito verificador
    
    If CStr(digito_11) <> (digito_rv) Then
        soma = 0
        soma = soma + Mid(REC, 1, 1) * 2
        soma = soma + Mid(REC, 2, 1) * 4
        soma = soma + Mid(REC, 3, 1) * 8
        
        resto = soma Mod 11         ' resto da divisão
        digito_11 = 11 - resto      ' digito verificador
        
        ' *****************************************************
        ' * Se o Cálculo for Igual a 10 ou 11, muda-se para 0 *
        ' *****************************************************
        If (digito_11 = 11) Or (digito_11 = 10) Then
            digito_11 = 0
        End If
   
        digito_rv = Mid(REC, 4, 1)  ' digito verificador
        
        If CStr(digito_11) <> (digito_rv) Then
           bOk = False              ' digito não confere
        End If
    End If

    VerificaReceita = bOk
End Function
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub

Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub



Function SalvaDARFPreto() As Boolean

  On Error GoTo ERRO_SALVADARFPRETO

  Dim RetAgencia As Integer
  Dim strEncripta   As String
  
  SalvaDARFPreto = False

  'Formatando o campo 'CÓDIGO DE REFERENCIA'
  TxtReferencia.Text = String(17 - Len(Trim(TxtReferencia.Text)), "0") & TxtReferencia.Text

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then
    'Verificar se a Agencia de Origem está OK
    If Not ValidaAgenciaPorDocto(Geral.Documento.Agencia, "", False) Then
        TxtVencimento.SetFocus
        Exit Function
    End If

    'Validar Código da Receita
    If Not ValidaReceita Then Exit Function

    'Verifica Dupla Digitação
    If ValidaDARF.sReferencia = "1" Then
      'Verificar se Referencia está zerada
      If Val(TxtReferencia.Text) = 0 Then
        MsgBox "Código de Referência inválido.", vbInformation, App.Title
        TxtReferencia.SetFocus
        Exit Function
      Else
        'Validar Referencia com 1 DV
        If Not VerificaReferenciaDV1(TxtReferencia.Text) Then
          If Not VerificaReferenciaDV2(TxtReferencia.Text) Then
            MsgBox "Código de Referência inválido.", vbInformation, App.Title
            TxtReferencia.SetFocus
            Exit Function
          End If
        End If
      End If
    Else
      'Não Validar DV , pedir confirmação caso referencia seja diferente de 0
      If Val(TxtReferencia.Text) <> 0 And AlteraValor = False Then
        If PicRef.Visible = False Then
          'Exibir tela para confirmação do campo 'REFERENCIA'
          Me.Height = 3740
          PicRef.Visible = True
          TxtReferencia2.SetFocus
          Exit Function
        Else
          'Verificar se Campos de Referencia conferem
          If Val(TxtReferencia.Text) <> Val(TxtReferencia2.Text) Then
            MsgBox "Referência não confere.", vbInformation + vbOKOnly, App.Title
            TxtReferencia2.SetFocus
            TxtReferencia2.SelStart = 0
            TxtReferencia2.SelLength = Len(TxtReferencia2.Text)
            Exit Function
          End If
        End If
      End If
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 16 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(16, CStr(Val(TxtCGCCPF.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVADARFPRETO

    'Atualizar / Inserir DARF Preto
    With qryAtualizaDARFPreto
      .rdoParameters(0) = Geral.DataProcessamento       'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto       'IdDocto
      .rdoParameters(2) = TxtReceita.Text               'CodigoReceita
      .rdoParameters(3) = TxtPeriodo.InverseText        'PeriodoApuracao
      .rdoParameters(4) = TxtCGCCPF.Text                'CGC / CPF
      .rdoParameters(5) = TxtReferencia.Text            'Referencia
      .rdoParameters(6) = TxtVencimento.InverseText     'Vencimento
      .rdoParameters(7) = Val(TxtPrincipal.Text) / 100  'Valor Principal
      .rdoParameters(8) = Val(TxtMulta.Text) / 100      'ValorMulta
      .rdoParameters(9) = Val(TxtJuros.Text) / 100      'Juros
      .rdoParameters(10) = Val(TxtValor.Text) / 100     'Valor
      .rdoParameters(11) = 16                           'TipoDocto
      .rdoParameters(12) = strEncripta                  'Autenticacao digital
      .Execute
    End With

    
    SalvaDARFPreto = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 16
  End If

  Exit Function

ERRO_SALVADARFPRETO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do DARF Preto.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PesquisaDARFPreto()

    On Error GoTo ERRO_PESQUISADARPRETO

    Dim sSql As String
    Dim RsDARFPreto As rdoResultset

    'Preencher os campos do DARF , caso encontre
    sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

    Set qryGetDARFPreto = Geral.Banco.CreateQuery("", "{call GetDARFPreto (" & sSql & ")}")

    Set RsDARFPreto = qryGetDARFPreto.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    If Not RsDARFPreto.EOF Then
        'Encontrou o DARF Preto -> Preencher os campos
        TxtReceita.Text = RsDARFPreto!CodigoReceita
        TxtPeriodo.Text = Mid(RsDARFPreto!PeriodoApuracao, 7, 2) & Mid(RsDARFPreto!PeriodoApuracao, 5, 2) & Mid(RsDARFPreto!PeriodoApuracao, 1, 4)
        TxtCGCCPF.Text = RsDARFPreto!CPFCGC
        TxtReferencia.Text = String(17 - Len(RsDARFPreto!Referencia), "0") & RsDARFPreto!Referencia
        TxtVencimento.Text = Mid(RsDARFPreto!vecto, 7, 2) & Mid(RsDARFPreto!vecto, 5, 2) & Mid(RsDARFPreto!vecto, 1, 4)
        TxtPrincipal.Text = Val(RsDARFPreto!ValorPrincipal * 100)
        TxtMulta.Text = Val(RsDARFPreto!ValorMulta * 100)
        TxtJuros.Text = Val(RsDARFPreto!Juros * 100)
        TxtValor.Text = Val(RsDARFPreto!Valor * 100)

        TxtPrincipal.SetFocus
    End If

    If AlteraValor = True Then
        'O Usuário só pode alterar os valores
        TxtReceita.Locked = True
        TxtPeriodo.Locked = True
        TxtCGCCPF.Locked = True
        TxtReferencia.Locked = True
        TxtVencimento.Locked = True
    
        TxtPrincipal.SetFocus
    End If

    Screen.MousePointer = vbDefault

    Exit Sub

ERRO_PESQUISADARPRETO:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Selecionar Dados do DARF Preto.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Function CamposOK() As Boolean

  On Error GoTo ERRO_CAMPOSOK

  Dim sData As String
  Dim sDataMov As String

  CamposOK = False

  'Período Apuração
  If Len(Trim(TxtPeriodo.Text)) = 0 Then
    MsgBox "Informe o Período de Apuração do Documento.", vbInformation, App.Title
    TxtPeriodo.SetFocus
    Exit Function
  End If

  'CGC / CPF
  If Len(Trim(TxtCGCCPF.Text)) = 0 Then
    MsgBox "Informe o CGC / CPF.", vbInformation, App.Title
    TxtCGCCPF.SetFocus
    Exit Function
  End If

  'Código da Receita
  If Len(Trim(TxtReceita.Text)) = 0 Then
    MsgBox "Informe o Código da Receita.", vbInformation, App.Title
    TxtReceita.SetFocus
    Exit Function
  End If

  'Data de Vencimento
  If Len(Trim(TxtVencimento.Text)) = 0 Then
    MsgBox "Informe a Data de Vencimento do Documento.", vbInformation, App.Title
    TxtVencimento.SetFocus
    Exit Function
  End If

  'Principal
  If Val(TxtPrincipal.Text) = 0 Then
    MsgBox "Informe o Valor do Principal.", vbInformation, App.Title
    TxtPrincipal.SetFocus
    Exit Function
  End If

  Call CalculaValorFinal

  'Preenchimento do Valor
  If Len(Trim(TxtValor.Text)) = 0 Then
    MsgBox "Informe o Valor Final do Documento.", vbInformation, App.Title
    TxtValor.SetFocus
    Exit Function
  End If

  'Validação do Valor
  If CCur(TxtValor.Text) < ValidaDARF.nValorLimite Then
    MsgBox "O Valor Total não pode ser menor que " & ValidaDARF.nValorLimite & ".", vbInformation, App.Title
    TxtPrincipal.SetFocus
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
  Set qryGetDARFPreto = Geral.Banco.CreateQuery("", "{? = call GetDARFPreto (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryAtualizaDARFPreto = Geral.Banco.CreateQuery("", "{call AtualizaDARFPreto (?,?,?,?,?,?,?,?,?,?,?,?,?)}")
  Set qryGetValidaDARFPreto = Geral.Banco.CreateQuery("", "{call GetValidaDARFPreto (?)}")

  Me.Height = 2820
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

  Set qryGetDARFPreto = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryAtualizaDARFPreto = Nothing
End Sub

Private Sub TxtCGCCPF_Change()
  If Len(Trim(TxtCGCCPF.Text)) = TxtCGCCPF.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtCGCCPF_GotFocus()
  TxtCGCCPF.SelStart = 0
  TxtCGCCPF.SelLength = Len(TxtCGCCPF.Text)
End Sub
Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtCGCCPF_LostFocus()
    '* Valida CGC *'
    If Len(Trim(TxtCGCCPF.Text)) = 0 Then Exit Sub
        If Not IsNumeric(TxtCGCCPF.Text) Then
            MsgBox "CGC incorreto, Redigite.", vbInformation, App.Title
            TxtCGCCPF.Text = ""
            TxtCGCCPF.SetFocus
            Exit Sub
        End If
End Sub
Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtJuros_LostFocus()
  Call CalculaValorFinal
End Sub
Private Sub txtMulta_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtMulta_LostFocus()
  Call CalculaValorFinal
End Sub
Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then
      KeyAscii = 0
      TxtPeriodo.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
      SendKeys "{TAB}"
  ElseIf KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtPrincipal_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtPrincipal_LostFocus()
  Call CalculaValorFinal
End Sub
Private Sub TxtReceita_Change()
  If Len(Trim(TxtReceita.Text)) = TxtReceita.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtReceita_GotFocus()
  TxtReceita.SelStart = 0
  TxtReceita.SelLength = Len(TxtReceita.Text)
End Sub
Private Sub txtReceita_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtReceita_LostFocus()
    '* Valida Receita *'
    If Len(Trim(TxtReceita.Text)) = 0 Then Exit Sub
        If Not IsNumeric(TxtReceita.Text) Then
            MsgBox "Código de Receita incorreto, Redigite.", vbInformation, App.Title
            TxtReceita.Text = ""
            TxtReceita.SetFocus
            Exit Sub
        End If
End Sub
Private Sub TxtReferencia_Change()
  If Len(Trim(TxtReferencia.Text)) = TxtReferencia.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtReferencia_GotFocus()
  TxtReferencia.SelStart = 0
  TxtReferencia.SelLength = Len(TxtReferencia.Text)
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtReferencia_LostFocus()
    '* Valida Referência *'
    If Len(Trim(TxtReferencia.Text)) = 0 Then Exit Sub
        If Not IsNumeric(TxtReferencia.Text) Then
            MsgBox "Código de Referência incorreto, Redigite.", vbInformation, App.Title
            TxtReferencia.Text = ""
            TxtReferencia.SetFocus
            Exit Sub
        End If
End Sub
Private Sub TxtReferencia2_GotFocus()
  TxtReferencia.SelStart = 0
  TxtReferencia.SelLength = Len(TxtReferencia.Text)
End Sub
Private Sub TxtReferencia2_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
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
Private Sub TxtVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace And TxtVencimento.Locked = False Then
      KeyAscii = 0
      TxtVencimento.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
      SendKeys "{TAB}"
  ElseIf KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Function CargaRegraValidaDARF() As Boolean

Dim rsGetRegra As rdoResultset
Dim qryGetRegra As rdoQuery

On Error GoTo Err_CargaRegraValidaDARF
    
    Erase RegraValidaDARF

    CargaRegraValidaDARF = False

    Set qryGetRegra = Geral.Banco.CreateQuery("", "{call GetRegraValidaDarfPreto }")

    Set rsGetRegra = qryGetRegra.OpenResultset(rdOpenStatic, rdConcurReadOnly)

    If Not rsGetRegra.EOF Then
        ReDim RegraValidaDARF(rsGetRegra.RowCount)

        While Not rsGetRegra.EOF
            RegraValidaDARF(rsGetRegra.AbsolutePosition).CodigoReceitaNaoAceito = rsGetRegra!CodigoReceitaNaoAceito
            RegraValidaDARF(rsGetRegra.AbsolutePosition).DataBase1 = rsGetRegra!DataBase1
            RegraValidaDARF(rsGetRegra.AbsolutePosition).DataBase2 = rsGetRegra!DataBase2
            'RegraValidaDARF(rsGetRegra.AbsolutePosition).ValorMinimoDocumento = rsGetRegra!ValorMinimoDocumento
            
            rsGetRegra.MoveNext
        Wend
        CargaRegraValidaDARF = True
    Else
        MsgBox "Não há parâmetros de Regras para validação de DARF PRETO." & vbCrLf & vbCrLf & "Favor contatar o suporte.", vbCritical + vbOKOnly, App.Title
    End If
    
Exit_CargaRegraValidaDARF:
    If Not (rsGetRegra Is Nothing) Then rsGetRegra.Close
    qryGetRegra.Close
    
    Exit Function

Err_CargaRegraValidaDARF:
    Screen.MousePointer = vbDefault
    Beep
    Select Case TratamentoErro("Erro na carga das Regras de DARF PRETO.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo Exit_CargaRegraValidaDARF

End Function


