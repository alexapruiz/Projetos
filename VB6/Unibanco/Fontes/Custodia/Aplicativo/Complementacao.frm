VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Complementacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Captura - Digitação"
   ClientHeight    =   5700
   ClientLeft      =   1530
   ClientTop       =   765
   ClientWidth     =   9330
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Complementacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5702.294
   ScaleMode       =   0  'User
   ScaleWidth      =   4704.81
   Begin VB.Frame FrameAviso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   2664
      TabIndex        =   0
      Top             =   2016
      Visible         =   0   'False
      Width           =   4305
      Begin VB.CommandButton CmdAviso 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1530
         TabIndex        =   1
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Atenção... Borderô com Diferenças"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   456
         TabIndex        =   2
         Top             =   300
         Width           =   3648
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "Complementacao.frx":000C
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.CommandButton CmdPainel 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1584
      Left            =   2532
      TabIndex        =   3
      Top             =   1968
      Visible         =   0   'False
      Width           =   4545
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   1  'Align Top
      Height          =   372
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9324
      _ExtentX        =   16457
      _ExtentY        =   661
      SimpleText      =   " "
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6991
            MinWidth        =   6991
            Picture         =   "Complementacao.frx":044E
            Text            =   "Número Boderô:  999999999999999999"
            TextSave        =   "Número Boderô:  999999999999999999"
            Object.ToolTipText     =   "Número do Borderô Corrente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5244
            MinWidth        =   5244
            Picture         =   "Complementacao.frx":0770
            Text            =   "Cheque / Posição:   11/20"
            TextSave        =   "Cheque / Posição:   11/20"
            Object.ToolTipText     =   "Quantidade de Cheques / Posição:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4544
            MinWidth        =   4544
            Picture         =   "Complementacao.frx":0A8A
            Text            =   "Data:    29/11/2000"
            TextSave        =   "Data:    29/11/2000"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Complementacao.frx":0DA4
   End
   Begin VB.Frame FrameCheque 
      Caption         =   "Cheques:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3852
      Left            =   2772
      TabIndex        =   13
      Top             =   408
      Width           =   6504
      Begin MSFlexGridLib.MSFlexGrid GridCheques 
         Height          =   3420
         Left            =   144
         TabIndex        =   25
         Top             =   252
         Width           =   6204
         _ExtentX        =   10927
         _ExtentY        =   6033
         _Version        =   393216
         Rows            =   12
         Cols            =   6
         FixedCols       =   0
         ForeColor       =   8388608
         ForeColorFixed  =   4210688
         GridColorFixed  =   128
         Redraw          =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "    Bco    |    Agência    |       Conta       |     Cheque     |              Valor           |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelNaoItens 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sem Itens Para Exibição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   144
         TabIndex        =   27
         Top             =   250
         Width           =   6204
      End
   End
   Begin VB.Frame FrameDataDeposito 
      Caption         =   "Borderô Datas: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3852
      Left            =   72
      TabIndex        =   14
      Top             =   408
      Width           =   2736
      Begin VB.ListBox ListData 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         ItemData        =   "Complementacao.frx":10BE
         Left            =   180
         List            =   "Complementacao.frx":10C0
         Style           =   1  'Checkbox
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Data cadastradas no Borderô Corrente"
         Top             =   564
         Width           =   2352
      End
      Begin VB.Label LabelDtDeposito 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datas de Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   250
         Width           =   2352
      End
   End
   Begin VB.Frame FrameDiferenca 
      Caption         =   "Diferenças:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5256
      Left            =   30
      TabIndex        =   6
      Top             =   384
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Timer TimerAtualiza 
         Enabled         =   0   'False
         Interval        =   50000
         Left            =   396
         Top             =   4644
      End
      Begin VB.TextBox TextDiferencas 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3588
         Left            =   144
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   8976
      End
      Begin VB.CommandButton CmdContinuar 
         Caption         =   "&Finalizar"
         Height          =   588
         Left            =   2952
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4500
         Width           =   1515
      End
      Begin VB.CommandButton CmdVoltar 
         Caption         =   "&Voltar"
         Height          =   588
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4500
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "           Data de Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   132
         TabIndex        =   12
         Top             =   324
         Width           =   3048
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         Qtde de Cheque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3192
         TabIndex        =   11
         Top             =   324
         Width           =   2772
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                    Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5988
         TabIndex        =   10
         Top             =   324
         Width           =   3120
      End
   End
   Begin VB.Frame FrameBotoesBordero 
      Caption         =   "Borderô:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1008
      Left            =   900
      TabIndex        =   21
      Top             =   4488
      Width           =   2712
      Begin VB.CommandButton CmdFinalizar 
         Caption         =   "&Finalizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1764
         Picture         =   "Complementacao.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   216
         Width           =   800
      End
      Begin VB.CommandButton CmdBordero 
         Caption         =   "No&vo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   144
         Picture         =   "Complementacao.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
      Begin VB.CommandButton CmdEnviarSupervisor 
         Caption         =   "Su&pervisor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   696
         Left            =   960
         Picture         =   "Complementacao.frx":16D6
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
   End
   Begin VB.Frame FrameBotoesCheque 
      Caption         =   "Cheques:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1008
      Left            =   3636
      TabIndex        =   16
      Top             =   4488
      Width           =   3564
      Begin VB.CommandButton CmdChequeAnterior 
         Caption         =   "&Anterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   180
         Picture         =   "Complementacao.frx":19E0
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
      Begin VB.CommandButton CmdChequePosterior 
         Caption         =   "Po&sterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   972
         Picture         =   "Complementacao.frx":20F6
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
      Begin VB.CommandButton CmdNewCheque 
         Caption         =   "&Novo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   1764
         Picture         =   "Complementacao.frx":2818
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   210
         Width           =   804
      End
      Begin VB.CommandButton CmdAltCheque 
         Caption         =   "Altera&r"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2568
         Picture         =   "Complementacao.frx":3352
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
   End
   Begin VB.Frame FrameBotoesFinal 
      Caption         =   "Encerrar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1008
      Left            =   7212
      TabIndex        =   4
      Top             =   4488
      Width           =   1044
      Begin VB.CommandButton CmdSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   696
         Left            =   144
         Picture         =   "Complementacao.frx":3E8C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   800
      End
   End
End
Attribute VB_Name = "Complementacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Controla eventos do list
 Dim ListEvento          As Boolean

'Classe das Consultas
 Private Proc_Sel        As New Custodia.Selecionar
 Private Proc_Atu        As New Custodia.Atualizar

'Dados do bordero ativo
 Private cIdBordero      As Long
 Private cNumBordero     As String * 19
 Private cDataDeposito   As Long
 
'Controle das datas do bordero ativo
 Dim VetDatas As Variant
 
'Controle de tempo para ativar (Timer)
 Private sTempo           As Integer

'Constantes
 Private Const STATUS_TXT_BORDERO = "Número Borderô :"
 Private Const STATUS_TXT_CHEQUE = "Cheque (Pos / Qtde) :"
 Private Const STATUS_TXT_DATA = "Data :"
 Private Const HABILITADO = &H80000005           'Window Background
 Private Const DESABILITADO = &H8000000F         'Button Face
 Private Const FRMCOMPLENORMAL = 6000
 Private Const FRMCOMPLECHEQUE = 8000
 Private Sub CmdAltCheque_Click()
    Dim Ret As enumRetornoModal
    Dim cIdCheque As Double
    Dim AtualDataDeposito
    
   'Guarda data p/ eventualidade de alteração da mesma pelo usuario
    AtualDataDeposito = cDataDeposito
    
   'Pega o Idcheque do cheque corrente do grid
    cIdCheque = GridCheques.TextMatrix(GridCheques.Row, 5)
    
   'Acerta o Form
    Complementacao.Height = FRMCOMPLECHEQUE
    
   'Passa os valores p/ o form de cheque
    Cheque.SetIdCheque cIdCheque
    Ret = Cheque.ShowModal(cIdBordero, cDataDeposito, cIdCheque, 5200)
    
   'Retorna o form a posicao normal
    Complementacao.Height = FRMCOMPLENORMAL
    
    If Ret = eRetornoOK Then
        'Verifica se foi alterado a data de deposito do cheque
        If cDataDeposito <> AtualDataDeposito Then
           'Caso a dt. alterada volta para a data corrente anterior e remonta o Grid
            cDataDeposito = AtualDataDeposito
            
            If CarregaCheques Then
                Navega True
            Else
                Navega False
            End If
        Else
           'Caso tenha alterado outras informções altera apenas a linha do mesmo
            If CarregaCheque(cIdCheque, False) Then
                Navega True
            Else
                Navega False
            End If
        End If
    End If

End Sub
Private Sub CmdAviso_Click()

    CmdPainel.Visible = False
    FrameAviso.Visible = False

End Sub
Private Sub CmdContinuar_Click()
    Dim Fechamento As New CalculoBordero       'Classe de calculo (fechamento)
          
    Fechamento.SetConnection g_cMainConnection
    Fechamento.IdBordero = cIdBordero
    Fechamento.DataProcessamento = Geral.DataProcessamento
    Fechamento.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
    Fechamento.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas

    If MsgBox("Tem Certeza que deseja encerrar borderô com pendências", vbInformation + vbYesNo) = vbYes Then
        If MudaStatus("4") Then
            'Executa Verificação de cheques inválidos
             Call Fechamento.VoltaStatusChequesIndevidos
             Call Fechamento.CalculaChequesIndevidosQTDE
             Call Fechamento.CalculaChequesIndevidosDATA
             
            'Inicializa variveis ao carregar o Form
             ExibeFinal True
             ListData.Clear
             cIdBordero = 0
             cNumBordero = 0
             Cabecalho False
             Botoes False, False
             Navega False
             
            'Desabilita Timer
             TimerAtualiza.Enabled = False
             
             CmdBordero.SetFocus
        End If
    End If
End Sub
Private Sub CmdEnviarSupervisor_Click()
   'Captura os dados do cheque corrente alimentando controle na Tela
    Dim Ret                 As Integer
    Dim Fechamento As New CalculoBordero       'Classe de calculo (fechamento)
          
    Fechamento.SetConnection g_cMainConnection
    Fechamento.IdBordero = cIdBordero
    Fechamento.DataProcessamento = Geral.DataProcessamento
    Fechamento.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
    Fechamento.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas
      
    If MsgBox("Confirma envio do Borderô para Supervisor", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If MudaStatus("5") Then
       'Executa Verificação de cheques inválidos
        Call Fechamento.VoltaStatusChequesIndevidos
        Call Fechamento.CalculaChequesIndevidosQTDE
        Call Fechamento.CalculaChequesIndevidosDATA
    
       'Inicializa variveis ao carregar o Form
        ListData.Clear
        cIdBordero = 0
        cNumBordero = 0
        Cabecalho False
        Botoes False, False
        Navega False
        
       'Desabilita Timer
        TimerAtualiza.Enabled = False

        CmdBordero.SetFocus
    End If
End Sub
Private Sub cmdFinalizar_Click()
    Dim sstr    As String
    Dim i       As Integer
    Dim Delay   As Date
    Dim Fechamento As New CalculoBordero       'Classe de calculo (fechamento)
            
    TextDiferencas.Text = ""
            
    Fechamento.SetConnection g_cMainConnection
    Fechamento.IdBordero = cIdBordero
    Fechamento.DataProcessamento = Geral.DataProcessamento
    
    If Not Fechamento.Calcula Then
       'Exibe msg por 1,5 segundos ou até usuário teclar enter
        CmdPainel.Visible = True
        FrameAviso.Visible = True
        DoEvents
        CmdAviso.SetFocus
        Delay = Time()
        
        While DateDiff("S", Delay, Time()) < 1.5
            DoEvents
        Wend
        
        CmdPainel.Visible = False
        FrameAviso.Visible = False
    
       'Exibe/Oculta controles
        ExibeFinal False
        sstr = ""
        For i = 1 To Fechamento.DataDeposito.Count
            If Fechamento.DataDeposito(i).DataDivergente Then
               sstr = sstr & Space(6)     '& vbTab
               sstr = sstr & FormataData(Fechamento.DataDeposito(i).DataDeposito, DD_MM_AAAA) & vbTab '& Space(5)
               'sStr = sStr & Format(Fechamento.DataDeposito(i).DiferencaQuantidade, "000")
               sstr = sstr & Space(5 - Len(Format(Fechamento.DataDeposito(i).DiferencaQuantidade, "000")))
               sstr = sstr & Format(Fechamento.DataDeposito(i).DiferencaQuantidade, "000")
               sstr = sstr & Space(27 - Len(Format(Fechamento.DataDeposito(i).DiferencaValor, MASK_VALOR)))
               sstr = sstr & Format(Fechamento.DataDeposito(i).DiferencaValor, MASK_VALOR) & vbCrLf
               TextDiferencas.Text = sstr
            End If
        Next
        
        CmdVoltar.SetFocus
    Else
    
        Fechamento.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
        Fechamento.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas
    
        Call Fechamento.VoltaStatusChequesIndevidos
        Call Fechamento.CalculaChequesIndevidosQTDE
        Call Fechamento.CalculaChequesIndevidosDATA
                
        If MudaStatus("R") Then
            'Exibe msg de OK
             MsgBox "Borderô Finalizado sem Pendências", vbInformation + vbOKOnly
             
            'Volta a tela normal
             ExibeFinal True
             
            'Reinicializa variveis iniciais
             ListData.Clear
             cIdBordero = 0
             cNumBordero = 0
             Botoes False, False
             Navega False
             
            'Desabilita Timer
             TimerAtualiza.Enabled = False
             
             CmdBordero.SetFocus
       End If
    End If
    
End Sub
Private Sub CmdBordero_Click()
   'Chamada da tela de Bordero inclusao/alteracao
    Dim Ret As enumRetornoModal
    Dim AntList As Integer
    Dim TEMP
     
    AntList = IIf(ListData.ListIndex >= 0, ListData.ListIndex, 0)
    
    Me.Hide
    Bordero.SetIdbordero cIdBordero
    
    'Erase VetDatas
    
   'Verifica retorno e variaveis
    Ret = Bordero.ShowModal(cIdBordero, cNumBordero, , VetDatas)
    
    If Ret = eRetornoOK Then
        If Not CarregaBordero(cIdBordero) Then
            Botoes False, False
            MsgBox "Erro"
            Exit Sub
        Else
           'Habilita Timer
            TimerAtualiza.Enabled = True
                                                
           'Habilita controles
            Cabecalho True
            Botoes True, False
            
           'Pre-seleciona o 1o. item do text
            SelecionaLista AntList
        End If
    End If
    
    Me.Show vbModal
    
End Sub
Private Sub CmdNewCheque_Click()
   'Inclusão de Novo Cheque
    Dim Ret As enumRetornoModal
    Dim cIdCheque As Double
    Dim UltimaDataBordero As Boolean
           
Pesquisa:
   'Se Cheques incluidos maior que digitado no bordero:
    If (GridCheques.Rows - 1) >= VerificaQtdeData(UltimaDataBordero) Then
        Complementacao.Height = FRMCOMPLENORMAL
        
       'Se for última data do list e do grid de datas do borderô
        If UltimaDataBordero Then
           'Para seleção para ultima data para inclusao de cheques
            Select Case MsgBox("Qtde de Cheques informados para esta data já incluída, Finalizar Borderô ?", vbYesNo + vbInformation, App.Title)
                Case vbYes
                   'Continuar (sem efeito)
                   cmdFinalizar_Click
                   Exit Sub
                Case vbNo
                    'Exit Sub
            End Select
        
        Else
           'Para seleção de nova data para inclusao de cheques
            Select Case MsgBox("Qtde de Cheques informados para esta data já incluída, Mudar p/ Próxima Data ?", vbYesNoCancel + vbInformation, App.Title)
                Case vbYes
                    SelecionaLista (ListData.ListIndex + 1)
                    GoTo Pesquisa
                Case vbNo
                   'Continuar (sem efeito)
                Case vbCancel
                    Exit Sub
            End Select
            
        End If
                    
    End If
        
    cIdCheque = 0  'Inclusao
    
   'Form abrigando form de cheque
    Complementacao.Height = FRMCOMPLECHEQUE
        
   'Chamada da Tela de Cheque
    Cheque.SetIdCheque cIdCheque
    
   'Verifica retorno da tela e variaveis
    Ret = Cheque.ShowModal(cIdBordero, cDataDeposito, cIdCheque, 5200)
            
    If Ret = eRetornoOK Then
                    
       'Carrega em Tela dados do novo cheque
        Call CarregaCheque(cIdCheque, True)
        
       'Habilita navegação entre cheques
        Navega True
        
       'Se retorno OK continua inclusão
        CmdNewCheque_Click
    Else
       'Form p/ tamanho normal
        Complementacao.Height = FRMCOMPLENORMAL
    End If
    
End Sub
Sub Navega(pNavega As Boolean)
   'Habilita controles de navegacao de cheques
    Dim sCheque As String
    
    sCheque = STATUS_TXT_CHEQUE & " "
    
    CmdAltCheque.Enabled = pNavega
    CmdChequeAnterior.Enabled = pNavega
    CmdChequePosterior.Enabled = pNavega
    
    If pNavega Then
        StatusBar.Panels(2).Text = sCheque & GridCheques.Row & " / " & GridCheques.Rows - 1
    Else
        StatusBar.Panels(2).Text = sCheque & "0 / 0"
    End If

End Sub
Sub Botoes(aBordero As Boolean, aCheque As Boolean)
   'Controle botoes de alteracao/inclusao de bordero e finalizar/sair
    If aBordero Then
        CmdBordero.Caption = "&Alterar"
        CmdFinalizar.Enabled = True
        CmdSair.Enabled = False
        CmdSair.Cancel = False
        CmdEnviarSupervisor.Enabled = True
    Else
        ListData.Visible = False
        GridCheques.Visible = False
        CmdBordero.Caption = "&Novo"
        CmdFinalizar.Enabled = False
        CmdSair.Enabled = True
        CmdSair.Cancel = True
        CmdEnviarSupervisor.Enabled = False
    End If
    
    CmdNewCheque.Enabled = aCheque
  
End Sub
Private Sub Cabecalho(ByVal pShow As Boolean)
   'Preenche a Barra de Status do Form
   
    Dim sBordero            As String
    
    sBordero = STATUS_TXT_BORDERO & " "
      
    If pShow Then
        sBordero = sBordero & cNumBordero
    End If
    
    ListData.Visible = CBool(pShow)
    GridCheques.Visible = CBool(pShow)
    
    StatusBar.Panels(1).Text = sBordero
    StatusBar.Panels(3).Text = STATUS_TXT_DATA & Format(Format(Geral.DataProcessamento, "0000/00/00"), "dd/mm/yyyy")

End Sub
Function CarregaBordero(pIdBordero As Long) As Boolean
   'Recebe datas de deposito do bordero (Busca no Banco)
    Dim rst                 As New ADODB.Recordset
        
    On Error GoTo Erro_CarregaBordero:
    
   'Limpa List
    ListData.Clear
       
    Set rst = g_cMainConnection.Execute(Proc_Sel.GetDatas(Geral.DataProcessamento, pIdBordero))
          
    Do While Not rst.EOF()
    
        ListData.AddItem FormataData(rst(0).Value, DD_MM_AAAA)
        rst.MoveNext
        
    Loop
        
    rst.Close
    CarregaBordero = True
    
    Exit Function
    
Erro_CarregaBordero:
    CarregaBordero = False
    Call TratamentoErro("Erro ao Carregar datas de depósito para o Bordero Corrente", Err)
    
End Function
Private Sub CmdVoltar_Click()
    ExibeFinal True
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    
   'Inicializa variveis ao carregar o Form
    cIdBordero = 0
    cNumBordero = 0
   
   'Barra de Status
    Cabecalho False
        
   'Desabilita Botoes
    Botoes False, False
    
   'Desabilta navegacao entre cheques
    Navega False
    
   'Inicializa scanner
    Call Principal.SetScanner
    
End Sub
Private Sub Form_Activate()

    If CmdNewCheque.Enabled Then
        CmdNewCheque.SetFocus
    Else
        If Complementacao.Height = FRMCOMPLENORMAL And CmdBordero.Enabled Then
            CmdBordero.SetFocus
        End If
    End If
   
 End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Finaliza Scanner
    Call Principal.DelScanner
   
End Sub
Private Sub GridCheques_DblClick()

    If GridCheques.Rows <= 1 Then
        Exit Sub
    Else
        CmdAltCheque_Click
    End If
    
End Sub
Private Sub GridCheques_GotFocus()

    If GridCheques.Rows <= 1 Then
        Exit Sub
    End If
    
    GridCheques.TopRow = GridCheques.Row
    
End Sub
Private Sub GridCheques_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then CmdAltCheque_Click
    
End Sub
Private Sub GridCheques_RowColChange()
   'New teste
   If GridCheques.Row <> 0 Then
        GridCheques.TopRow = GridCheques.Row
   End If
  
   If GridCheques.Rows > 12 Then
       GridCheques.Width = 6400
   Else
       GridCheques.Width = 6200
   End If
    
   'Acerta foco p/ linha inteira
    GridCheques.Col = 0
    GridCheques.ColSel = 4
    
    Navega True

End Sub
Private Sub CmdChequeAnterior_Click()
    With GridCheques
        If .Row = 1 Then
            Exit Sub
        Else
            .Row = .Row - 1
            .SetFocus
        End If
    End With
End Sub
Private Sub CmdChequePosterior_Click()
    With GridCheques
        If .Row = .Rows - 1 Then
            Exit Sub
        Else
            .Row = .Row + 1
            .SetFocus
        End If
    End With
End Sub
Private Sub ListData_Click()

   'Selecionar e checar Item selecionado/clicado
    Dim i   As Integer
    
   'Controla evento do List
    If ListEvento Then
        Exit Sub
    End If
    
    ListEvento = True
    SelecionaLista ListData.ListIndex
    ListEvento = False
    
   'Habilita controles
    Cabecalho True
    Botoes True, True
    
   'Verifica data deposito selecionada/clicada atualizando variavel da data corrente
    For i = 0 To ListData.ListCount - 1
        If ListData.Selected(i) Then

            cDataDeposito = Format(ListData.Text, "YYYYMMDD")

            If CarregaCheques Then
                Navega True
            Else
                Navega False
            End If
           
            Exit Sub
        End If
    Next i
    
End Sub
Private Sub ListData_ItemCheck(Item As Integer)

   'Controla evento do List
    If ListEvento Then
        Exit Sub
    Else
        ListEvento = True
        SelecionaLista Item
        ListEvento = False
    End If

End Sub
Private Sub SelecionaLista(ByVal pItem As Integer)
   'Controla evento do List
   
    Dim i As Integer
    
    For i = 0 To ListData.ListCount - 1
        ListData.Selected(i) = False
    Next i
    
    ListData.Selected(pItem) = True

End Sub
Sub ExibeFinal(pHabilita As Boolean)
   'Tela com demonstração das diferenças
   
    FrameBotoesBordero.Enabled = pHabilita
    FrameBotoesCheque.Enabled = pHabilita
    FrameBotoesFinal.Enabled = pHabilita
    FrameCheque.Visible = pHabilita
    FrameDataDeposito.Visible = pHabilita
    FrameDiferenca.Visible = IIf(pHabilita, False, True)
    
End Sub
Function MudaStatus(pStatus As String) As Boolean
   'Muda status do bordero p/ Supervisor
    Dim Ret         As Integer
    Dim Repete      As Boolean
    
    Repete = True
    
    On Error GoTo Errostatus:
        
Repetir:
    Call g_cMainConnection.Execute(Proc_Atu.AtualizaStatusBordero(Geral.DataProcessamento, cIdBordero, pStatus), Ret, adCmdText)
    
    If Ret <> 0 Then
        MudaStatus = True
        Exit Function
    End If
    
Errostatus:
    MudaStatus = False
    Call TratamentoErro("Erro ao Mudar Status do Borderô", Err, Repete)
    If Repete Then Resume Repetir
End Function
Function CarregaCheques() As Boolean
   'Procura por todos os cheques cadastrados na data, inclui no Grid
    Dim rst As New ADODB.Recordset
    Dim i   As Integer
    Dim id  As String
        
    On Error GoTo ErroCheques:
    
    Set rst = g_cMainConnection.Execute(Proc_Sel.GetChequesComplementados(Geral.DataProcessamento, cIdBordero, cDataDeposito))
  
    If rst.EOF Then
       'Se data selecionada ñ ha cheques retorna falso
        GridCheques.HighLight = flexHighlightNever
        GridCheques.Rows = 1
        CarregaCheques = False
    Else
        i = 1
        GridCheques.HighLight = flexHighlightAlways
        GridCheques.Rows = rst.RecordCount + 1
        
        While (Not rst.EOF)
            With rst
                GridCheques.TextMatrix(i, 0) = rst(0)
                GridCheques.TextMatrix(i, 1) = rst(1)
                GridCheques.TextMatrix(i, 2) = rst(2) & "-" & rst(3)
                GridCheques.TextMatrix(i, 3) = rst(4)
                GridCheques.TextMatrix(i, 4) = Format(rst(5), MASK_VALOR) & Space(5)
                GridCheques.TextMatrix(i, 5) = rst(6)
                rst.MoveNext
            End With
            i = i + 1
        Wend
        
        GridCheques.Row = 1
        CarregaCheques = True
   End If
       
   GridCheques_RowColChange
      
   Exit Function

ErroCheques:
    Call TratamentoErro("Erro ao Selecionar Cheques", Err)

End Function
Function CarregaCheque(pIdCheque, Inclui As Boolean) As Boolean
   'Procura dados do cheque Alterado ou Incluido
    Dim rst As New ADODB.Recordset
    Dim id  As String
    
    On Error GoTo ErroCheque:
    
    Set rst = g_cMainConnection.Execute(Proc_Sel.GetChequesComplementados(Geral.DataProcessamento, cIdBordero, , pIdCheque))
       
    If rst.EOF Then
       'Se data selecionada ñ ha cheques retorna falso
        CarregaCheque = False
        Err.Raise 997, App.Title, "Erro ao Carregar Cheque Complementado"
    Else
        If Inclui Then
            GridCheques.Rows = GridCheques.Rows + 1
            GridCheques.Row = GridCheques.Rows - 1
        End If
        
        GridCheques.SetFocus
        GridCheques.HighLight = flexHighlightAlways
        GridCheques_RowColChange
        
        GridCheques.TextMatrix(GridCheques.Row, 0) = rst(0)
        GridCheques.TextMatrix(GridCheques.Row, 1) = rst(1)
        GridCheques.TextMatrix(GridCheques.Row, 2) = rst(2) & "-" & rst(3)
        GridCheques.TextMatrix(GridCheques.Row, 3) = rst(4)
        GridCheques.TextMatrix(GridCheques.Row, 4) = Format(rst(5), MASK_VALOR) & Space(5)
        GridCheques.TextMatrix(GridCheques.Row, 5) = rst(6)
       
        CarregaCheque = True
   End If
   
   Exit Function

ErroCheque:
    Call TratamentoErro("Erro ao Selecionar Cheques", Err)
End Function
Private Sub timerAtualiza_Timer()
On Error GoTo Erro:
    Dim rsHoraAtual As ADODB.Recordset

    TimerAtualiza.Enabled = False
    sTempo = sTempo + Int(TimerAtualiza.Interval / 1000)

    If sTempo + Int(TimerAtualiza.Interval / 1000) >= g_Parametros.TMP_Pendente Then

         'Obs.: Utiliza a tabela StatusBordero apenas para obter a hora atual do servidor
         Set rsHoraAtual = g_cMainConnection.Execute("select distinct time() from PARAMETRO")

         'Atualizar a hora do bordero
         Call g_cMainConnection.Execute(Proc_Atu.AtualizaHoraAtualBordero(Geral.DataProcessamento, cIdBordero, rsHoraAtual(0)))

         sTempo = 0
    End If
    
    TimerAtualiza.Enabled = True
    Set rsHoraAtual = Nothing
    
Exit Sub

Erro:
    Call TratamentoErro("Falha na Atualização da Hora", Err, False, False)
End Sub
Function VerificaQtdeData(pUltimaDataBordero) As Integer
   'Varre Vetor, Retorna Qtde cheques digitados no bordero e se é última data do list
    Dim i As Integer

    For i = 0 To UBound(VetDatas)
        If VetDatas(i, 0) = ListData.Text Then
           'Achar a Data Selecionada e verificar a qtde de cheques correspondente
            VerificaQtdeData = VetDatas(i, 1)
            
           'Retorna True se for Ultima data incluida no bordero
            If i = UBound(VetDatas) Then
                pUltimaDataBordero = True
            Else
                pUltimaDataBordero = False
            End If
            
            Exit For
        End If
    Next i
   
End Function
