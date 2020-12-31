VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form MensagemErro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2745
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtMensErro 
      Height          =   1704
      Left            =   72
      TabIndex        =   5
      Top             =   432
      Width           =   7716
      _ExtentX        =   13600
      _ExtentY        =   3016
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"MensagemErro.frx":0000
   End
   Begin VB.CommandButton CmdHelpScan 
      Caption         =   "&Ajuda"
      Height          =   500
      Left            =   6912
      Picture         =   "MensagemErro.frx":0082
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2200
      Width           =   850
   End
   Begin VB.CommandButton CmdRepetir 
      Caption         =   "&Repetir"
      Height          =   500
      Left            =   2484
      Picture         =   "MensagemErro.frx":038C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2200
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.CommandButton cmdDescricao 
      Caption         =   "&Descrição"
      Height          =   500
      Left            =   4344
      Picture         =   "MensagemErro.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2200
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   500
      Left            =   3420
      Picture         =   "MensagemErro.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2200
      Width           =   850
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Picture         =   "MensagemErro.frx":0CAA
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Tratamento de Erro"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   120
      TabIndex        =   2
      Top             =   96
      Width           =   7620
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   24
      Shape           =   4  'Rounded Rectangle
      Top             =   24
      Width           =   7812
   End
End
Attribute VB_Name = "MensagemErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTexto          As String           'Texto  do Programador
Dim lError          As Variant          'Código de Erro do Objeto AdoError
Dim sDescricao      As String           'Descricao do Erro do Objeto AdoError
Dim ExibeCmdRepetir As Boolean          'Exibe Botão repetir se necessário
Dim ExibeCmdHelpScanner As Boolean

Enum EnumCores
    nulo = 0
    AZUL = vbBlue
    Verde = vbGreen
    Vermelho = vbRed
    Preto = vbBlack
    Amarelo = vbYellow
    Rosa = vbMagenta
    Cian = vbCyan
End Enum

Public Sub ShowModal(ByVal psTexto As String _
                   , ByVal plError As Variant _
                   , ByVal psDescricao As String _
                   , Optional ByRef psRepetir As Boolean _
                   , Optional ByRef psScanHelp As Boolean)
                   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       * Recebe parametro da Função TratamentoErro *                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sTexto = psTexto
    lError = plError
    sDescricao = psDescricao
    CmdRepetir.Visible = CBool(psRepetir)
    CmdHelpScan.Visible = CBool(psScanHelp)
    
  
    
    MensagemErro.Show vbModal
    
    psRepetir = CBool(ExibeCmdRepetir)

End Sub
Private Sub cmdDescricao_Click()
    Static Estado As Boolean

    txtMensErro.Text = ""
        
    If Estado Then
        txtMensErro.Text = sTexto
        cmdDescricao.Caption = "&Descrição"
        Estado = False
         
    Else
        txtMensErro.Text = lError & " - " & sDescricao
        cmdDescricao.Caption = "&Erro"
        Estado = True
    End If
    
   'Formatação
    FormataText txtMensErro, 0, Len(txtMensErro), , vbCenter, 10, , True, False
End Sub
Private Sub CmdHelpScan_Click()
    Dim Texto As String
    
    txtMensErro.Visible = True
    txtMensErro.Text = ""
    txtMensErro.SelAlignment = vbLeftJustify
           
    Texto = vbTab & vbTab & UCase("Solução de Problemas com Scanner:") & vbCrLf & vbCrLf
    
    Texto = Texto & "INICIALIZAÇÃO:" & vbCrLf
    Texto = Texto & Chr(149) & " Verifique se os Parâmetros estão configurados corretamente." & vbCrLf
    Texto = Texto & Chr(149) & " O cabo deve estar na mesma porta serial da estação, especificada em Parâmetro." & vbCrLf
    Texto = Texto & Chr(149) & " Se scanner = LA93 o mesmo deve ser ligado antes da estação." & vbCrLf
    
    Texto = Texto & "LEITURA:" & vbCrLf
    Texto = Texto & Chr(149) & " Os cabos do Scanner estão devidamente conectados(Energia e/ou Serial)." & vbCrLf
    Texto = Texto & Chr(149) & " O cabo serial está conectado na mesma porta especificada no módulo"
    Texto = Texto & " parâmetro do sistema." & vbCrLf
    Texto = Texto & Chr(149) & " Scanner configurado no sistema é o mesmo que esta sendo utilizado." & vbCrLf
    
    Texto = Texto & "GRAVAÇÃO E OBTENÇÃO DE PARAMETROS:" & vbCrLf
    Texto = Texto & Chr(149) & " Se scanner = [L100] e ambiente = NT, deve existir arquivo DTC32NT.DLL no diretorio de sistema: [c:\Win...\System32]." & vbCrLf
    Texto = Texto & Chr(149) & " Se scanner = [L100] e ambiente = 9X, deve existir arquivo DTC329X.DLL no diretorio de sistema: [c:\Win...\System]." & vbCrLf
    Texto = Texto & Chr(149) & " Se scanner = [LA93] em qualquer ambiente, deve existir arquivo LA93.DLL no diretorio de sistema e o driver VipsDrv, "
    Texto = Texto & "instalado na estação." & vbCrLf
    
    txtMensErro.Text = Texto
       
    FormataText txtMensErro, 0, 35, , , 12, Vermelho, True, True
    FormataText txtMensErro, 36, 18, , , 10, AZUL, True, False
    FormataText txtMensErro, 54, 210, , , 10, , False, False
    FormataText txtMensErro, 263, 10, , , 10, AZUL, True, False
    FormataText txtMensErro, 272, 236, , , 10, , False, False
    FormataText txtMensErro, 507, 36, , , 10, AZUL, True, False
    FormataText txtMensErro, 544, 400, , , 10, , False, False
End Sub
Private Sub cmdSair_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   * Verifica Tipo de Tratamento: Inclusão ou Alteração *                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ExibeCmdRepetir = False
      
    Unload Me
    
End Sub
Private Sub Form_Activate()
   txtMensErro.Text = ""
   txtMensErro.Text = sTexto
    
  'Formatação
   FormataText txtMensErro, 0, Len(txtMensErro), , vbCenter, 10, , True, False

End Sub
Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       * Mostra Texto que Programador Informou *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Me.Top = 1000
    Me.Left = 2100
End Sub
Private Sub cmdRepetir_click()
    
    ExibeCmdRepetir = True
    Unload Me
        
End Sub
Sub FormataText(pObjeto As Object, pSelIni As Long, pSellen As Long, Optional pNomeFonte As String, _
                Optional pAlinhamento As AlignmentConstants, Optional pFontSize As Long, Optional pFontColor As EnumCores, _
                Optional pNegrito As Boolean, Optional PGrifado As Boolean)
                

    pObjeto.SelStart = pSelIni
    pObjeto.SelLength = pSellen
    pObjeto.SetFocus
    
    pObjeto.SelFontName = IIf(pNomeFonte = "", "Courrier New", pNomeFonte)
    pObjeto.SelAlignment = IIf(pAlinhamento = 0, vbLeftJustify, pAlinhamento)
    pObjeto.SelFontSize = IIf(pFontSize = 0, 10, pFontSize)
    pObjeto.SelColor = IIf(pFontColor = nulo, vbBlack, pFontColor)
    pObjeto.SelBold = CBool(pNegrito)
    pObjeto.SelUnderline = CBool(PGrifado)
    
    pObjeto.SelStart = 0
    pObjeto.SelLength = 0
        
End Sub


