VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form FichaCompensacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Fichas de Compensação"
   ClientHeight    =   2520
   ClientLeft      =   204
   ClientTop       =   2580
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   11652
   Begin VB.Frame fraBanco 
      Height          =   1452
      Left            =   3912
      TabIndex        =   35
      Top             =   -696
      Visible         =   0   'False
      Width           =   5892
      Begin VB.CommandButton cmdCancelaBanco 
         Caption         =   "Cancela&r"
         Height          =   300
         Left            =   4560
         TabIndex        =   39
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox txtBanco 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   38
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código do Banco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   720
         TabIndex        =   37
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label lblBanco 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Confirmação do Código de Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   0
         TabIndex        =   36
         Top             =   120
         Width           =   5892
      End
   End
   Begin VB.PictureBox PicRef 
      Height          =   888
      Left            =   120
      ScaleHeight     =   840
      ScaleWidth      =   10740
      TabIndex        =   33
      Top             =   2436
      Visible         =   0   'False
      Width           =   10788
      Begin VB.Frame Frame1 
         Caption         =   "Confirmação do Código de Barras"
         Height          =   756
         Left            =   48
         TabIndex        =   34
         Top             =   24
         Width           =   8928
         Begin VB.TextBox TxtCodigo4C 
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
            Left            =   5928
            MaxLength       =   1
            TabIndex        =   22
            Top             =   264
            Width           =   360
         End
         Begin VB.TextBox TxtCodigo5C 
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
            Left            =   6792
            MaxLength       =   14
            TabIndex        =   23
            Top             =   264
            Width           =   1776
         End
         Begin VB.TextBox TxtCodigo1C 
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
            Left            =   312
            MaxLength       =   10
            TabIndex        =   19
            Top             =   264
            Width           =   1308
         End
         Begin VB.TextBox TxtCodigo3C 
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
            Left            =   3996
            MaxLength       =   11
            TabIndex        =   21
            Top             =   264
            Width           =   1428
         End
         Begin VB.TextBox TxtCodigo2C 
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
            Left            =   2076
            MaxLength       =   11
            TabIndex        =   20
            Top             =   264
            Width           =   1404
         End
      End
      Begin VB.CommandButton CmdCancelarRef 
         Caption         =   "&Cancelar"
         Height          =   324
         Left            =   9288
         TabIndex        =   24
         Top             =   264
         Width           =   1128
      End
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   9600
      Picture         =   "FichaCompensacao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   144
      Width           =   864
   End
   Begin DATEEDITLib.DateEdit TxtVencimento 
      Height          =   372
      Left            =   9672
      TabIndex        =   5
      Top             =   1284
      Width           =   1572
      _Version        =   65537
      _ExtentX        =   2773
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
      Left            =   1968
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1284
      Width           =   1404
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
      Left            =   3876
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1284
      Width           =   1428
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
      Left            =   144
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1284
      Width           =   1308
   End
   Begin VB.TextBox txtCodigo5 
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
      Left            =   6672
      MaxLength       =   14
      TabIndex        =   4
      Top             =   1284
      Width           =   1776
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
      Left            =   5820
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1284
      Width           =   360
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   696
      Left            =   5232
      Picture         =   "FichaCompensacao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   144
      Width           =   864
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   696
      Left            =   6108
      Picture         =   "FichaCompensacao.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   144
      Width           =   864
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   696
      Left            =   6984
      Picture         =   "FichaCompensacao.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   144
      Width           =   864
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      Height          =   696
      Left            =   7860
      Picture         =   "FichaCompensacao.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   144
      Width           =   864
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      Height          =   696
      Left            =   8736
      Picture         =   "FichaCompensacao.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   144
      Width           =   864
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   696
      Left            =   10476
      Picture         =   "FichaCompensacao.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   144
      Width           =   864
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtAbatimento 
      Height          =   372
      Left            =   3948
      TabIndex        =   8
      Top             =   1980
      Width           =   1884
      _Version        =   65537
      _ExtentX        =   3323
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
   Begin CURRENCYEDITLib.CurrencyEdit txtDesconto 
      Height          =   372
      Left            =   2064
      TabIndex        =   7
      Top             =   1980
      Width           =   1836
      _Version        =   65537
      _ExtentX        =   3238
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
      Left            =   5892
      TabIndex        =   9
      Top             =   1980
      Width           =   1836
      _Version        =   65537
      _ExtentX        =   3238
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtValorBase 
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   1980
      Width           =   1884
      _Version        =   65537
      _ExtentX        =   3323
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
      Height          =   372
      Left            =   9648
      TabIndex        =   11
      Top             =   1980
      Width           =   1884
      _Version        =   65537
      _ExtentX        =   3323
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
   Begin CURRENCYEDITLib.CurrencyEdit txtMora 
      Height          =   372
      Left            =   7776
      TabIndex        =   10
      Top             =   1980
      Width           =   1836
      _Version        =   65537
      _ExtentX        =   3238
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "( + ) Mora"
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
      Left            =   7704
      TabIndex        =   40
      Top             =   1728
      Width           =   864
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Linha Digitável do Código de Barras"
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
      Height          =   264
      Left            =   168
      TabIndex        =   32
      Top             =   960
      Width           =   3180
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   516
      Picture         =   "FichaCompensacao.frx":1546
      Top             =   252
      Width           =   384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Fichas de Compensação"
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
      TabIndex        =   31
      Top             =   372
      Width           =   2196
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "( = ) Valor Cobrado"
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
      Left            =   9708
      TabIndex        =   30
      Top             =   1716
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Base"
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
      Left            =   156
      TabIndex        =   29
      Top             =   1716
      Width           =   984
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "( + ) Juros"
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
      Left            =   5952
      TabIndex        =   28
      Top             =   1716
      Width           =   912
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "( - ) Desconto"
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
      Left            =   2076
      TabIndex        =   27
      Top             =   1716
      Width           =   1224
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "( - ) Abatimento"
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
      Left            =   4008
      TabIndex        =   26
      Top             =   1716
      Width           =   1368
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   9768
      TabIndex        =   25
      Top             =   984
      Width           =   1056
   End
End
Attribute VB_Name = "FichaCompensacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryGetFichaCompensacao          As rdoQuery
Private qryRemoveTipoDocumento          As rdoQuery
Private qryAtualizaDocumentoExcluido    As rdoQuery
Private qryAtualizaFichaCompensacao     As rdoQuery
'Private qryGetAgenf                     As rdoQuery
Private qryGetFichaBarDuplicada         As rdoQuery

'Declaração de Variáveis de trabalho
Private bFormating                      As Boolean
Private mForm                           As Form
Private bActivate                       As Boolean
Private ObjCtrl                         As Object
Private bAlterouBarras                  As Boolean

Public Alterou                          As Boolean
Public AlteraValor                      As Boolean
Private bDigitacaoBanco                 As Boolean


Sub AjustesIniciais()

  'Setando as Variáveis do RDO
  Set qryGetFichaCompensacao = Geral.Banco.CreateQuery("", "{? = call GetFichaCompensacao (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaFichaCompensacao = Geral.Banco.CreateQuery("", "{? = call AtualizaFichaCompensacao (?,?,?,?,?,?,?,?,?,?,?,?)}")
'  Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (?)}")
  Set qryGetFichaBarDuplicada = Geral.Banco.CreateQuery("", "{? = call GetFichaBarDuplicada (?,?,?,?,?)}")
End Sub
Sub CalculaValorCobrado()

  On Error GoTo ERRO_CALCULAVALORCOBRADO

  Dim Valor As Currency

  'Verificar se foi informado o Valor Base
  If Val(TxtValorBase.Text) = 0 Then
    TxtValor.Text = ""
    Exit Sub
  End If

  'Verificar se foi informado Juros
  If Val(TxtJuros.Text) <> 0 Then
    Valor = Val(TxtValorBase.Text) + Val(TxtJuros.Text)
  Else
    Valor = TxtValorBase.Text
  End If

  'Verificar se foi informado Descontos
  If Val(txtDesconto.Text) <> 0 Then
    Valor = Valor - Val(txtDesconto.Text)
  End If

  'Verificar se foi informado Abatimentos
  If Val(TxtAbatimento.Text) <> 0 Then
    Valor = Valor - Val(TxtAbatimento.Text)
  End If

  'Verificar se foi informado Mora
  If Val(txtMora.Text) <> 0 Then
    Valor = Valor + Val(txtMora.Text)
  End If
  
  'Transportar o Valor Final para a tela
  TxtValor.Text = Valor

  Exit Sub

ERRO_CALCULAVALORCOBRADO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Cobrado.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function DefineTipoDocto(ByVal sCodigo As String) As Integer

    On Error GoTo ERRO_DEFINETIPODOCTO

    'Determinar o Tipo do Documento de acordo com o código de barras
    If (Mid(sCodigo, 1, 1) <> "8") Then
        'Ficha de Compensacao
        If (Mid(sCodigo, 1, 3) = "409") Then

            If (Mid(sCodigo, 20, 2) = "04") Then
                DefineTipoDocto = 28           'UNICOBRANÇA
                Exit Function
            End If

            If (Mid(sCodigo, 20, 1) = "6") Then
                DefineTipoDocto = 29           'COBRANÇA IMEDIATA UNIBANCO
                Exit Function
            End If

            If (Val((Mid(sCodigo, 20, 1))) >= 1) And (Val((Mid(sCodigo, 20, 1))) <= 5) Then
                DefineTipoDocto = 30           '1,2,3,4,5 COBRANÇA ESPECIAL UNIBANCO
                Exit Function
            End If

            If (Val((Mid(sCodigo, 20, 1))) > 6) Then
                DefineTipoDocto = 0            'NÃO TEM TIPO DEFINIDO
            End If
        Else
            DefineTipoDocto = 31              'COBRANÇA DE TERCEIROS
        End If
    End If

    Exit Function

ERRO_DEFINETIPODOCTO:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Verificar o Tipo de Documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function

Private Sub cmdCancelaBanco_Click()

    txtBanco = ""
    fraBanco.Visible = False
    bDigitacaoBanco = False
    
End Sub

Private Sub CmdCancelarRef_Click()
  Me.Height = 2820
  PicRef.Visible = False
  TxtCodigo1C.Text = ""
  TxtCodigo2C.Text = ""
  TxtCodigo3C.Text = ""
  TxtCodigo4C.Text = ""
  TxtCodigo5C.Text = ""
End Sub
Private Sub cmdConfirmar_Click()

    If fraBanco.Visible Then
        txtBanco.SetFocus
        Exit Sub
    End If

    If SalvaFicha Then
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
    bDigitacaoBanco = False
    fraBanco.Visible = False
    fraBanco.Left = 120
    fraBanco.Top = 840
    Alterou = False
    
    Call AjustesIniciais

    Call PesquisaFicha
    
    'Se form de complementação, força digitação do código de barras
    If LCase(mForm.Name) = "complementacao" Then
        bAlterouBarras = True
    Else
        bAlterouBarras = False
    End If
    
    bActivate = False
End Sub
Public Function CalculaDataCodigoBarras(pvnIndice As Integer) As String
    Dim nDia As Integer
    Dim nMes As Integer
    Dim nAno As Integer
    Dim nInd As Integer
    Dim nUltimoDia As Integer

    If pvnIndice < 1001 Then
        CalculaDataCodigoBarras = "00000000"
        Exit Function
    End If

    CalculaDataCodigoBarras = Format(DateAdd("d", pvnIndice - 1001, "04/07/2000"), "yyyymmdd")

End Function
Private Function DVModulo11(ByVal pvsNumero As String) As Integer

    Dim nSoma As Integer        ' Somatória dos Elementos do Código de Barras
    Dim nPeso As Integer        ' Peso para somar elementos do Código de Barras
    Dim nInd As Integer         ' Variável auxiliar

    nSoma = 0
    nPeso = 9

    For nInd = Len(Trim(pvsNumero)) To 1 Step -1
      nPeso = IIf(nPeso <> 9, nPeso + 1, 2)
      nSoma = nSoma + (Val(Mid(pvsNumero, nInd, 1)) * nPeso)
      DoEvents
    Next nInd

    DVModulo11 = IIf((11 - (nSoma Mod 11)) > 9, 1, 11 - (nSoma Mod 11))
End Function
Function SalvaFicha() As Boolean

  On Error GoTo ERRO_SALVAFICHA

  Dim sCodigoBarras     As String
  Dim TipoDocto         As Integer
  Dim RetAgencia        As Integer
  Dim CodigoErrado      As Boolean
  Dim strEncripta       As String
  
  SalvaFicha = False
  Alterou = False

    'Verificar se todos os campos estão preenchidos
    If CamposOK Then
        'Validar Primeiro Campo do Código de Barras
        If Not Modulo10(txtCodigo1.Text, 10) Then
            'MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            CodigoErrado = True
            If txtCodigo1.Enabled = True Then
                txtCodigo1.SetFocus
            End If
        End If

        'Validar Segundo Campo do Código de Barras
        If Not Modulo10(txtCodigo2.Text, 11) Then
            'MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            CodigoErrado = True
            If txtCodigo2.Enabled = True Then
                txtCodigo2.SetFocus
            End If
        End If

        'Validar Terceiro Campo do Código de Barras
        If Not Modulo10(txtCodigo3.Text, 11) Then
            'MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            CodigoErrado = True
            If txtCodigo3.Enabled = True Then
                txtCodigo3.SetFocus
            End If
        End If

        'Validar Quarto Campo do Código de Barras
        If Len(Trim(txtCodigo4.Text)) = 0 Then
            'MsgBox "Informe o Quarto campo do Código de Barras.", vbInformation, App.Title
            CodigoErrado = True
            If txtCodigo4.Enabled = True Then
                txtCodigo4.SetFocus
            End If
        End If

        'Formatar o Campo de Valor do Código de Barras
        txtCodigo5.Text = Format(txtCodigo5.Text, String(14, "0"))

        'Verificar se o Código de Barras está no novo formato
        If Val(Mid(txtCodigo5.Text, 1, 4)) >= 1000 Then
            'Novo Formato -> Validar Super DV do Código de Barras
            If Val(txtCodigo4.Text) <> DVModulo11(Mid(txtCodigo1.Text, 1, 4) & txtCodigo5.Text & Mid(txtCodigo1.Text, 5, 5) & Mid(txtCodigo2.Text, 1, 10) & Mid(txtCodigo3.Text, 1, 10)) Then
                CodigoErrado = True
                If txtCodigo1.Enabled = True Then
                    txtCodigo1.SetFocus
                End If
            End If
        End If

        'Verificar se o documento é do Unibanco e possui codigo de barras zerado
        If Val(txtCodigo2.Text) = 0 Then
            If Mid(txtCodigo1.Text, 1, 3) = "409" Then
                MsgBox "Código de Barras Inválido.", vbInformation, App.Title

                If txtCodigo2.Enabled = True Then
                    txtCodigo2.SetFocus
                End If
                Exit Function
            Else
                CodigoErrado = True
            End If
        End If

        If Val(txtCodigo3.Text) = 0 Then
            If Mid(txtCodigo1.Text, 1, 3) = "409" Then
                MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    
                If txtCodigo3.Enabled = True Then
                    txtCodigo3.SetFocus
                End If
    
                Exit Function
            Else
                CodigoErrado = True
            End If
        End If

        'Exibir Mensagem de erro , se houver
        If Mid(txtCodigo1.Text, 1, 3) = "409" Then
            'Ficha do Unibanco -> Exibir Mensagem
            If CodigoErrado Then
                MsgBox "Código de Barras Inválido.", vbInformation, App.Title
                Exit Function
            End If
        Else
            If CodigoErrado Then
                If PicRef.Visible = False Then
                    'Exibir tela para confirmação dos campos do codigo de barras
                    Me.Height = 3740
                    PicRef.Visible = True
                    TxtCodigo1C.SetFocus
                    Exit Function
                Else
                    'Verificar primeiro campo do codigo de barras
                    If Val(txtCodigo1.Text) <> Val(TxtCodigo1C.Text) Then
                        MsgBox "Código de Barras não confere.", vbInformation, App.Title
                        TxtCodigo1C.SetFocus
                        TxtCodigo1C.SelStart = 0
                        TxtCodigo1C.SelLength = Len(TxtCodigo1C.Text)
                        Exit Function
                    End If
    
                    'Verificar segundo campo do codigo de barras
                    If Val(txtCodigo2.Text) <> Val(TxtCodigo2C.Text) Then
                        MsgBox "Código de Barras não confere.", vbInformation, App.Title
                        TxtCodigo2C.SetFocus
                        TxtCodigo2C.SelStart = 0
                        TxtCodigo2C.SelLength = Len(TxtCodigo2C.Text)
                        Exit Function
                    End If
    
                    'Verificar terceiro campo do codigo de barras
                    If Val(txtCodigo3.Text) <> Val(TxtCodigo3C.Text) Then
                        MsgBox "Código de Barras não confere.", vbInformation, App.Title
                        TxtCodigo3C.SetFocus
                        TxtCodigo3C.SelStart = 0
                        TxtCodigo3C.SelLength = Len(TxtCodigo3C.Text)
                        Exit Function
                    End If
    
                    'Verificar quarto campo do codigo de barras
                    If Val(txtCodigo4.Text) <> Val(TxtCodigo4C.Text) Then
                        MsgBox "Código de Barras não confere.", vbInformation, App.Title
                        TxtCodigo4C.SetFocus
                        TxtCodigo4C.SelStart = 0
                        TxtCodigo4C.SelLength = Len(TxtCodigo4C.Text)
                        Exit Function
                    End If
    
                    'Verificar quinto campo do codigo de barras
                    If Val(txtCodigo5.Text) <> Val(TxtCodigo5C.Text) Then
                        MsgBox "Código de Barras não confere.", vbInformation, App.Title
                        TxtCodigo5C.SetFocus
                        TxtCodigo5C.SelStart = 0
                        TxtCodigo5C.SelLength = Len(TxtCodigo5C.Text)
                        Exit Function
                    End If
                End If
            End If
        End If

        'Solicita digitação do código do banco
        If Not bDigitacaoBanco Then
            fraBanco.Visible = True
            txtBanco.SetFocus
            Exit Function
        End If
        bDigitacaoBanco = False

        'Definindo o Codigo de Barras para gravar na tabela
        sCodigoBarras = Left(Trim(txtCodigo1.Text), 4) & Trim(txtCodigo4.Text) & _
                        Trim(txtCodigo5.Text) & Mid(Trim(txtCodigo1.Text), 5, 5) & _
                        Left(Trim(txtCodigo2.Text), 10) & Left(Trim(txtCodigo3.Text), 10)

        If Len(sCodigoBarras) <> 44 Then
            Beep
            MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            txtCodigo1.SetFocus
            Exit Function
        End If

        'Definindo o Tipo do documento , de acordo com o código de barras
        TipoDocto = DefineTipoDocto(sCodigoBarras)

        If TipoDocto = 0 Then
            MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            txtCodigo1.SetFocus
            Exit Function
        End If

        'Validar Agencia de Origem da Capa
        If Geral.Capa.agefsstmovi <> 0 Then
            RetAgencia = ValidaAgenciaFicha
        Else
            RetAgencia = ValidaAgencia(Geral.Documento.Agencia, TxtVencimento.Text, True)
        End If

        'Verificar Retorno da Função
        Select Case RetAgencia
        Case 1
            'Documento Vencido
            If Geral.Capa.IdEnv_Mal = "E" Then
                'Envelope com titulo de outros bancos -> Pedir confirmação antes de devolver
                If MsgBox("Este documento está vencido. Confirma ? ", vbInformation + vbYesNo, App.Title) = vbNo Then
                    TxtVencimento.SetFocus
                    Exit Function
                Else
                    If Left(sCodigoBarras, 3) <> "409" And Left(sCodigoBarras, 3) <> "230" Then
                        'Devolver documento
                        With qryAtualizaDocumentoExcluido
                            .rdoParameters(0).Direction = rdParamOutput
                            .rdoParameters(1).Value = Geral.DataProcessamento
                            .rdoParameters(2).Value = Geral.Documento.IdDocto
                            .rdoParameters(3).Value = "D"       'Status
                            .rdoParameters(4).Value = 0         'Duplicidade
                            .rdoParameters(5).Value = 208       'Ocorrencia
                            .Execute

                            Geral.Documento.Status = "D"

                            If .rdoParameters(0).Value <> 0 Then
                                GoTo ERRO_SALVAFICHA
                            End If
                        End With
                    End If
                End If
            ElseIf Geral.Capa.IdEnv_Mal = "M" Then
                'Para o Novo Malote não permitir documento vencido de outro banco
                If (Left(CStr(Geral.Capa.Num_Malote), 1) = "9") And _
                    (Len(Trim(Geral.Capa.Num_Malote)) = 11) Then

                    If MsgBox("Este documento está vencido. Confirma ? ", vbInformation + vbYesNo, App.Title) = vbNo Then
                        TxtVencimento.SetFocus
                        Exit Function
                    Else
                        If Left(sCodigoBarras, 3) <> "409" And Left(sCodigoBarras, 3) <> "230" Then
                            With qryAtualizaDocumentoExcluido
                                .rdoParameters(0).Direction = rdParamOutput
                                .rdoParameters(1).Value = Geral.DataProcessamento
                                .rdoParameters(2).Value = Geral.Documento.IdDocto
                                .rdoParameters(3).Value = "D"       'Status
                                .rdoParameters(4).Value = 0         'Duplicidade
                                .rdoParameters(5).Value = 208       'Ocorrencia
                                .Execute

                                Geral.Documento.Status = "D"

                                If .rdoParameters(0).Value <> 0 Then
                                    GoTo ERRO_SALVAFICHA
                                End If
                            End With
                        End If
                    End If
                Else
                    If MsgBox("Este documento pertence a um Malote e está vencido. Confirma ?", vbYesNo + vbInformation, App.Title) = vbNo Then
                        'Malote Antigo -> Pedir Confirmação
                        TxtVencimento.SetFocus
                        Exit Function
                    End If
                End If
            Else
                'Tipo Indefinido
                MsgBox "Não foi possível definir se o documento pertence a um Envelope ou Malote " & Chr(13) & _
                "para aplicar regra de validação de Data de Vencimento.", vbInformation, App.Title
                Exit Function
            End If
        Case 2
            'Agencia em Feriado
            MsgBox "A agência de origem está em feriado.", vbInformation, App.Title
            TxtVencimento.SetFocus
            Exit Function
        Case 3
            'Agencia Fechada
            MsgBox "A agência de origem está fechada.", vbInformation, App.Title
            TxtVencimento.SetFocus
            Exit Function
        Case 4
            'Agencia não Cadastrada
            MsgBox "A agência de origem não está cadastrada.", vbInformation, App.Title
            TxtVencimento.SetFocus
            Exit Function
        End Select

        'Verificar se já existe outro documento na base com o mesmo codigo de barras _
        desconsiderando o valor (igual Arrecadacao Eletronica)

        With qryGetFichaBarDuplicada
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1).Value = Geral.DataProcessamento
            .rdoParameters(2).Value = Mid(sCodigoBarras, 1, 4)
            .rdoParameters(3).Value = Mid(sCodigoBarras, 5, 15)
            .rdoParameters(4).Value = Mid(sCodigoBarras, 20, 25)
            .rdoParameters(5).Value = Geral.Documento.IdDocto
            .Execute
        End With

        If qryGetFichaBarDuplicada(0).Value = 1 Then
            'Encontrou outro documento com código de barras igual e valor diferente
            MsgBox "Já existe outro documento com o mesmo código de barras, por isso, este documento deve ser complementado como 'Arrecadação Convencional';", vbInformation, App.Title
            txtCodigo1.SetFocus
            Exit Function
        End If

        'Verificar se o Documento pertence à outro Tipo
        If Geral.Documento.TipoDocto <> 10 And Geral.Documento.TipoDocto <> 28 And _
            Geral.Documento.TipoDocto <> 29 And Geral.Documento.TipoDocto <> 30 And _
            Geral.Documento.TipoDocto <> 31 And Geral.Documento.TipoDocto <> 0 Then
            With qryRemoveTipoDocumento
                .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
                .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
                .Execute
            End With
        End If

        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(TipoDocto, sCodigoBarras)
        If strEncripta = "" Then GoTo ERRO_SALVAFICHA

        'Atualizar / Inserir Ficha
        With qryAtualizaFichaCompensacao
            .rdoParameters(0).Direction = rdParamReturnValue            'Parametro de Retorno
            .rdoParameters(1) = Geral.DataProcessamento                 'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto                 'IdDocto
            .rdoParameters(3) = sCodigoBarras                           'Codigo de Barras
            .rdoParameters(4) = TxtVencimento.InverseText               'Vencimento
            .rdoParameters(5) = Val(TxtValorBase.Text) / 100            'Valor Base
            .rdoParameters(6) = Val(TxtJuros.Text) / 100                'Juros
            .rdoParameters(7) = Val(txtDesconto.Text) / 100             'Desconto
            .rdoParameters(8) = Val(TxtAbatimento.Text) / 100           'Abatimento
            .rdoParameters(9) = Val(TxtValor.Text) / 100                'Valor
            .rdoParameters(10) = TipoDocto                              'TipoDocto
            .rdoParameters(11) = Val(txtMora.Text) / 100                'Mora
            .rdoParameters(12) = strEncripta                            'Autenticacao digital
            .Execute
        End With
    
        SalvaFicha = True
        Alterou = True
        
        'Atualizar o Controle Global
        Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
        Geral.Documento.Leitura = sCodigoBarras
        Geral.Documento.TipoDocto = TipoDocto
    End If

    Exit Function

ERRO_SALVAFICHA:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar Dados da Ficha de Compensação.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function
Function CamposOK() As Boolean

CamposOK = False

  'Primeiro Campo do Código de Barras
  If Len(Trim(txtCodigo1.Text)) <> 10 Or Val(txtCodigo1.Text) = 0 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    If txtCodigo1.Enabled = True Then
      txtCodigo1.SetFocus
    End If
    Exit Function
  End If

  'Segundo Campo do Código de Barras
  If Len(Trim(txtCodigo2.Text)) <> 11 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    If txtCodigo2.Enabled = True Then
      txtCodigo2.SetFocus
    End If
    Exit Function
  End If

  'Terceiro Campo do Código de Barras
  If Len(Trim(txtCodigo3.Text)) <> 11 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    If txtCodigo3.Enabled = True Then
      txtCodigo3.SetFocus
    End If
    Exit Function
  End If

  'Quinto Campo do Código de Barras
  If Len(Trim(txtCodigo5.Text)) <> 14 Then
    MsgBox "Código de Barras Inválido.", vbInformation, App.Title
    If txtCodigo5.Enabled = True Then
      txtCodigo5.SetFocus
    End If
    Exit Function
  End If

  'Vencimento
  If Len(Trim(TxtVencimento.Text)) = 0 Then
    MsgBox "Informe a Data de Vencimento do Documento.", vbInformation, App.Title
    If TxtVencimento.Enabled = True Then
      TxtVencimento.SetFocus
    End If
    Exit Function
  End If

  'Valor Base
  If Val(TxtValorBase.Text) = 0 Then
    MsgBox "Informe o Valor Base da Ficha.", vbInformation, App.Title
    TxtValorBase.SetFocus
    Exit Function
  End If

  'Calcula o Valor Cobrado
  Call CalculaValorCobrado

  'Valor
  If Val(TxtValor.Text) = 0 Then
    MsgBox "Não é permitido gravar documentos com valor cobrado zerado.", vbInformation, App.Title
    TxtValor.SetFocus
    Exit Function
  End If

  If CCur(TxtValor.Text) < 0 Then
    MsgBox "O Valor Cobrado não pode ser negativo. Verifique os valores digitados.", vbInformation, App.Title
    TxtValorBase.SetFocus
    Exit Function
  End If

  'Validar o Código do Banco
  If Not ValidaCodigoBanco(Mid(txtCodigo1.Text, 1, 3)) Then
    MsgBox "Código de Banco não participante do Sistema de Compensação.", vbInformation, App.Title
    txtCodigo1.SetFocus
    Exit Function
  End If

  CamposOK = True
  
End Function
Sub PesquisaFicha()

    On Error GoTo ERRO_PESQUISAFICHA

    Dim sSql As String
    Dim RsFicha As rdoResultset
    Dim CBInvalido As Boolean

    If Len(Trim(Geral.Documento.Leitura)) <> 0 Then
        txtCodigo1.Text = Mid(Geral.Documento.Leitura, 1, 4) + Mid(Geral.Documento.Leitura, 20, 5)
        txtCodigo2.Text = Mid(Geral.Documento.Leitura, 25, 10)
        txtCodigo3.Text = Mid(Geral.Documento.Leitura, 35, 10)
        txtCodigo4.Text = Mid(Geral.Documento.Leitura, 5, 1)
        txtCodigo5.Text = Mid(Geral.Documento.Leitura, 6, 14)

        If Me.AlteraValor = False Then
            'Verificar se o super DV bate
            If Val(txtCodigo4.Text) = DVModulo11(Mid(txtCodigo1.Text, 1, 4) & txtCodigo5.Text & Mid(txtCodigo1.Text, 5, 5) & Mid(txtCodigo2.Text, 1, 10) & Mid(txtCodigo3.Text, 1, 10)) Then
                'Super DV OK
                txtCodigo1.Text = txtCodigo1.Text & Format(DV10(txtCodigo1.Text), "0")
                txtCodigo2.Text = txtCodigo2.Text & Format(DV10(txtCodigo2.Text), "0")
                txtCodigo3.Text = txtCodigo3.Text & Format(DV10(txtCodigo3.Text), "0")
            End If
        Else
            txtCodigo1.Text = txtCodigo1.Text & Format(DV10(txtCodigo1.Text), "0")
            txtCodigo2.Text = txtCodigo2.Text & Format(DV10(txtCodigo2.Text), "0")
            txtCodigo3.Text = txtCodigo3.Text & Format(DV10(txtCodigo3.Text), "0")
        End If

        'Verificar se os tres campos estão preenchidos
        If Len(Trim(txtCodigo1.Text)) = 10 And InStr(txtCodigo1.Text, " ") = 0 _
            And Len(Trim(txtCodigo2.Text)) = 11 And InStr(txtCodigo2.Text, " ") = 0 _
            And Len(Trim(txtCodigo3.Text)) = 11 And InStr(txtCodigo3.Text, " ") = 0 Then

            'Validar Primeiro Campo do Código de Barras
            If Modulo10(txtCodigo1.Text, 10) Then
                'Validar Segundo Campo do Código de Barras
                If Modulo10(txtCodigo2.Text, 11) Then
                    'Validar Terceiro Campo do Código de Barras
                    If Modulo10(txtCodigo3.Text, 11) Then
                        'Código de Barras OK
                        TxtVencimento.SetFocus
                    End If
                End If
            End If
        Else
            CBInvalido = True
        End If
    Else
        CBInvalido = True
    End If

    'Preencher os campos da Ficha , caso encontre
    sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

    Set qryGetFichaCompensacao = Geral.Banco.CreateQuery("", "{call GetFichaCompensacao (" & sSql & ")}")

    Set RsFicha = qryGetFichaCompensacao.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    If Not RsFicha.EOF Then
        'Encontrou a Ficha -> Preencher os campos
        TxtVencimento.Text = Format(DataDDMMAAAA(RsFicha!vecto), "00000000")
        TxtValorBase.Text = RsFicha!ValorBase * 100
        TxtJuros.Text = RsFicha!Juros * 100
        txtDesconto.Text = RsFicha!Desconto * 100
        TxtAbatimento.Text = RsFicha!Abatimento * 100
        TxtValor.Text = RsFicha!Valor * 100
        txtMora.Text = RsFicha!Mora * 100

        TxtVencimento.SelStart = 0
        TxtVencimento.SelLength = Len(TxtVencimento.Text)
    End If

    If AlteraValor = True Then
        'O Usuário só pode alterar os campos de valor
        txtCodigo1.Locked = True
        txtCodigo2.Locked = True
        txtCodigo3.Locked = True
        txtCodigo4.Locked = True
        txtCodigo5.Locked = True
        TxtVencimento.Locked = True

        TxtValorBase.SetFocus
    End If

    If CBInvalido = True Then
        txtCodigo1.SetFocus
    ElseIf AlteraValor = True Then
        TxtValorBase.SetFocus
    Else
        TxtVencimento.SetFocus
    End If

    Screen.MousePointer = vbDefault

    Exit Sub

ERRO_PESQUISAFICHA:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Selecionar Dados da Ficha.", Err, rdoErrors)
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

    cmdSair.CausesValidation = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = 0 Then Alterou = False

  Set qryGetFichaCompensacao = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryAtualizaDocumentoExcluido = Nothing
  Set qryAtualizaFichaCompensacao = Nothing
  
End Sub

Private Sub TxtAbatimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub TxtAbatimento_LostFocus()
  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub

Private Sub txtBanco_GotFocus()
    
    txtBanco.SelStart = 0
    txtBanco.SelLength = txtBanco.MaxLength
    
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyReturn Xor KeyAscii = 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtBanco.Text)) <> 3 Then
            Beep
            MsgBox "Favor digitar os 3 dígitos do código de Banco", vbCritical, App.Title
            Exit Sub
        End If
        If txtBanco <> Left(txtCodigo1.Text, 3) Then
            Beep
            MsgBox "Código do banco não confere com o documento, Favor verificar!", vbCritical, App.Title
            fraBanco.Visible = False
            txtCodigo1.SetFocus
            Exit Sub
        End If

        bDigitacaoBanco = True
        fraBanco.Visible = False
        Call cmdConfirmar_Click
    End If
    
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
    
    bAlterouBarras = True
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub
Private Sub txtCodigo1_Validate(Cancel As Boolean)

   If Len(Trim(txtCodigo1.Text)) > 0 And Val(txtCodigo1.Text) <> 0 And Not bActivate Then
      'Validar Primeiro Campo do Código de Barras
      If Not Modulo10(Format(txtCodigo1.Text, String(txtCodigo1.MaxLength, "0")), 10) Then
         MsgBox "Código de Barras Inválido.", vbInformation, App.Title
         txtCodigo1.SelStart = 0
         txtCodigo1.SelLength = Len(txtCodigo1.Text)
         Cancel = True
      End If
   End If
End Sub
Private Sub TxtCodigo1C_Change()
  If Len(Trim(TxtCodigo1C.Text)) = TxtCodigo1C.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub TxtCodigo1C_GotFocus()
  TxtCodigo1C.SelStart = 0
  TxtCodigo1C.SelLength = TxtCodigo1C.MaxLength
End Sub
Private Sub TxtCodigo1C_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtCodigo1C_LostFocus()
'* Valida Código de Barras 1º campo *'
'* Código de Barras digitado tem que ser igual ao informado*'

    Set ObjCtrl = Screen.ActiveControl
    
    
    If Not ObjCtrl Is Nothing Then
        If (LCase(ObjCtrl.Name) = LCase("CmdCancelarRef")) Or _
           (LCase(ObjCtrl.Name) = LCase("CmdSair")) Then Exit Sub
    End If
    
    If Len(Trim(TxtCodigo1C.Text)) = 0 Then Exit Sub

    If txtCodigo1 <> TxtCodigo1C And Trim(TxtCodigo1C) <> "" Then
        MsgBox "Código digitado não pode ser diferente do informado acima.", vbInformation, App.Title
        SelecionarTexto TxtCodigo1C
        Exit Sub
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
Private Sub txtCodigo2_Validate(Cancel As Boolean)
   If Len(Trim(txtCodigo2.Text)) > 0 And Val(txtCodigo2.Text) <> 0 And Not bActivate Then
      'Validar Segundo Campo do Código de Barras
      If Not Modulo10(Format(txtCodigo2.Text, String(txtCodigo2.MaxLength, "0")), 11) Then
         MsgBox "Código de Barras Inválido.", vbInformation, App.Title
         txtCodigo2.SelStart = 0
         txtCodigo2.SelLength = Len(txtCodigo2.Text)
         Cancel = True
      End If
   End If
End Sub
Private Sub TxtCodigo2C_Change()
  If Len(Trim(TxtCodigo2C.Text)) = TxtCodigo2C.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub TxtCodigo2C_GotFocus()
  TxtCodigo2C.SelStart = 0
  TxtCodigo2C.SelLength = TxtCodigo2C.MaxLength
End Sub
Private Sub TxtCodigo2C_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtCodigo2C_LostFocus()
'* Valida Codigo de Barras - 2º campo *'
'* Código de Barras digitado tem que ser igual ao informado*'

    Set ObjCtrl = Screen.ActiveControl
    
    If Not ObjCtrl Is Nothing Then
        If (LCase(ObjCtrl.Name) = LCase("CmdCancelarRef")) Or _
           (LCase(ObjCtrl.Name) = LCase("CmdSair")) Then Exit Sub
    End If
    
    If Len(Trim(TxtCodigo2C.Text)) = 0 Then Exit Sub

    If txtCodigo2 <> TxtCodigo2C And Trim(TxtCodigo2C) <> "" Then
        MsgBox "Código digitado não pode ser diferente do informado acima.", vbInformation, App.Title
        SelecionarTexto TxtCodigo2C
        Exit Sub
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
Private Sub txtCodigo3C_LostFocus()
'* Valida Codigo de Barras - 3º campo *'
'* Código de Barras digitado tem que ser igual ao informado*'
    Set ObjCtrl = Screen.ActiveControl
    
    
    If Not ObjCtrl Is Nothing Then
        If (LCase(ObjCtrl.Name) = LCase("CmdCancelarRef")) Or _
           (LCase(ObjCtrl.Name) = LCase("CmdSair")) Then Exit Sub
    End If
    
    If Len(Trim(TxtCodigo3C.Text)) = 0 Then Exit Sub

    If txtCodigo3 <> TxtCodigo3C Then
        MsgBox "Código digitado não pode ser diferente do informado acima.", vbInformation, App.Title
        SelecionarTexto TxtCodigo3C
        Exit Sub
    End If

End Sub
Private Sub txtCodigo3_Validate(Cancel As Boolean)
   If Len(Trim(txtCodigo3.Text)) > 0 And Val(txtCodigo3.Text) <> 0 And Not bActivate Then
      'Validar Terceiro Campo do Código de Barras
      If Not Modulo10(Format(txtCodigo3.Text, String(txtCodigo3.MaxLength, "0")), 11) Then
         MsgBox "Código de Barras Inválido.", vbInformation, App.Title
         txtCodigo3.SelStart = 0
         txtCodigo3.SelLength = Len(txtCodigo3.Text)
         Cancel = True
      End If
   End If
End Sub
Private Sub TxtCodigo3C_Change()
  If Len(Trim(TxtCodigo3C.Text)) = TxtCodigo3C.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub TxtCodigo3C_GotFocus()
  TxtCodigo3C.SelStart = 0
  TxtCodigo3C.SelLength = TxtCodigo3C.MaxLength
End Sub
Private Sub TxtCodigo3C_KeyPress(KeyAscii As Integer)
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

Private Sub txtCodigo4C_LostFocus()
'* Valida Codigo de Barras - 4º campo *'
'* Código de Barras digitado tem que ser igual ao informado*'
    Set ObjCtrl = Screen.ActiveControl
    
    
    If Not ObjCtrl Is Nothing Then
        If (LCase(ObjCtrl.Name) = LCase("CmdCancelarRef")) Or _
           (LCase(ObjCtrl.Name) = LCase("CmdSair")) Then Exit Sub
    End If
    
    If Len(Trim(TxtCodigo4C.Text)) = 0 Then Exit Sub

    If txtCodigo4 <> TxtCodigo4C Then
        MsgBox "Código digitado não pode ser diferente do informado acima.", vbInformation, App.Title
        TxtCodigo4C.SelStart = 0
        TxtCodigo4C.SelLength = (TxtCodigo4C.MaxLength)
        TxtCodigo4C.SetFocus
        Exit Sub
    End If

End Sub
Private Sub txtCodigo4_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub
Private Sub TxtCodigo4C_Change()
  If Len(Trim(TxtCodigo4C.Text)) = TxtCodigo4C.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub TxtCodigo4C_GotFocus()
  TxtCodigo4C.SelStart = 0
  TxtCodigo4C.SelLength = TxtCodigo4C.MaxLength
End Sub
Private Sub TxtCodigo4C_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtCodigo5_Change()
  If bFormating Then Exit Sub
  If Len(Trim(txtCodigo5.Text)) = txtCodigo5.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub txtCodigo5_GotFocus()
  txtCodigo5.SelStart = 0
  txtCodigo5.SelLength = txtCodigo5.MaxLength
End Sub
Private Sub txtCodigo5_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub
Private Sub txtCodigo5_LostFocus()

    Dim sData As String

    'Formatar o Campo Codigo5
    bFormating = True
    txtCodigo5.Text = Format(Val(txtCodigo5.Text), String(14, "0"))
    bFormating = False

    'Se o documento for do formato antigo -> Codigo5 = Valor Base
    If Val(txtCodigo5.Text) <> 0 Then
        If Val(Mid(txtCodigo5.Text, 1, 1)) > 0 Then
            TxtValorBase.Text = Mid(txtCodigo5.Text, 5)
        Else
            TxtValorBase.Text = txtCodigo5.Text
        End If
    End If

    'Calcular Data de Vencimento de acordo com o código de barras (Formulário Novo)
    If Mid(txtCodigo5.Text, 1, 4) > 1000 Then
        sData = CalculaDataCodigoBarras(Mid(txtCodigo5.Text, 1, 4))
        TxtVencimento.Text = Mid(sData, 7, 2) & Mid(sData, 5, 2) & Mid(sData, 1, 4)
    End If
End Sub
Private Sub TxtCodigo5C_GotFocus()
  TxtCodigo5C.SelStart = 0
  TxtCodigo5C.SelLength = TxtCodigo5C.MaxLength
End Sub
Private Sub TxtCodigo5C_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    If ValidaCampo5 = False Then Exit Sub
    Call cmdConfirmar_Click
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDesconto_LostFocus()
  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub
Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtJuros_LostFocus()
  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub

Private Sub txtMora_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMora_LostFocus()
  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
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
Private Sub TxtValorBase_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub TxtValorBase_LostFocus()
  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub
Private Sub TxtVencimento_GotFocus()
  TxtVencimento.SelStart = 0
  TxtVencimento.SelLength = Len(TxtVencimento.Text)
End Sub
Private Sub TxtVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace And TxtVencimento.Locked = False Then
      KeyAscii = 0
      TxtVencimento.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
      SendKeys "{TAB}"
  ElseIf KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf KeyAscii = 42 Then
    If ValidaDataVencto = False Then Exit Sub
    cmdConfirmar_Click
  End If
  
End Sub
Function ValidaCampo5() As Boolean

    'Código de Barras digitado tem que ser igual ao informado
    If Len(Trim(TxtCodigo5C.Text)) = 0 Then Exit Function

    If Format(txtCodigo5, String(txtCodigo5.MaxLength, "0")) <> Format(txtCodigo5, String(TxtCodigo5C.MaxLength, "0")) Then
        MsgBox "Código digitado não pode ser diferente do informado acima.", vbInformation, App.Title
        TxtCodigo5C.SelStart = 0
        TxtCodigo5C.SelLength = (TxtCodigo5C.MaxLength)
        TxtCodigo5C.SetFocus
        ValidaCampo5 = False
        Exit Function
    End If
    ValidaCampo5 = True
End Function
Function ValidaDataVencto() As Boolean
'* Valida Data de Vencimento *'

    If Not IsDate(TxtVencimento.Text) Then
        ValidaDataVencto = False
        SendKeys "{Tab}"
    Else
        ValidaDataVencto = True
    End If
    
End Function
Function ValidaAgenciaFicha() As Integer

  'Código de Retorno
  '0 - Data de Vencimento OK
  '1 - Documento Vencido
  '2 - Agencia em Feriado
  '3 - Agencia Fechada
  '4 - Agencia não cadastrada
  '5 - Data não Verificada

  ValidaAgenciaFicha = 5

    'Verificar o Status
    If Geral.Capa.agefsstmovi = 9 Then
        'Feriado
        ValidaAgenciaFicha = 2
        Exit Function

    ElseIf Geral.Capa.agefsstmovi = 0 Then
        'Agencia Fechada
        ValidaAgenciaFicha = 3
        Exit Function

    ElseIf Geral.Capa.agefsstmovi = 2 Then
        'Agencia Aberta -> Verificar data do Movimento Anterior
        If DataAAAAMMDD(TxtVencimento.Text) <= TransformaDataAAAAMMDD(CStr(Geral.Capa.agefsdtmvan)) Then
            'A Data de Vencimento é menor ou igual à data do Movimento Anterior -> Não Aceitar
            ValidaAgenciaFicha = 1
            Exit Function
        End If
    End If

    ValidaAgenciaFicha = 0
End Function

