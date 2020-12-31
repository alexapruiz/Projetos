VERSION 5.00
Begin VB.Form frmComunica 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta e Log"
   ClientHeight    =   2580
   ClientLeft      =   2295
   ClientTop       =   2115
   ClientWidth     =   5070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerEnter 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   180
      Top             =   2070
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   4965
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "frmComunica.frx":0000
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data do Servidor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   7
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Multi-Agência:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   6
         Top             =   855
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número da Estação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   5
         Top             =   1350
         Width           =   2145
      End
      Begin VB.Label LabelNoEstacao 
         AutoSize        =   -1  'True
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   3240
         TabIndex        =   4
         Top             =   1344
         Width           =   480
      End
      Begin VB.Label LabelMultiAgencia 
         AutoSize        =   -1  'True
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   252
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   480
      End
      Begin VB.Label LabelDataServidor 
         AutoSize        =   -1  'True
         Caption         =   "99/99/9999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdComunica 
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
      Height          =   420
      Left            =   1845
      TabIndex        =   0
      Top             =   2025
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Esc - Encerra."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4005
      TabIndex        =   8
      Top             =   2340
      Width           =   1005
   End
End
Attribute VB_Name = "frmComunica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComunica_Click()

    
   'Desabilita timer de entrada automatica
    TimerEnter.Enabled = False
           
    Call GetCaixa
    Call IniOpcoes
    
   '''''''''''''''''''''''''''''''''''''''''
   ' Chama form para começar a comunicação '
   '''''''''''''''''''''''''''''''''''''''''
    
    Me.Hide
    
    frmShow.TimerComunica.Interval = 1000
    frmShow.Show vbModal
    
    Exit Sub
    
TrataErro:
    
    Select Case TratamentoErro("Não foi possível fazer a inicialização do Robô.", Err)
        Case eSair
            End
        Case eRepetir
           Resume
        Case eContinuar
            Resume Next
    End Select
        

   
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
Private Sub Form_Load()
    
On Error GoTo TrataErro
    
    Dim ValorAlcada         As String
    Dim AgenciaUBB          As Integer
    Dim PrzVencimento       As String
    Dim RstUBB              As Recordset
    Dim RstMDI              As Recordset
    Dim CargaTabelasOK       As String * 1
   
   'variaveis do header
'    Set RstMDI = MDIQuery.getContraPartida(Geral.rstDocto!Evento)
'
'    If RstMDI.EOF Then
'        RstMDI.Close
'        MsgBox "ATENÇÃO !!! (31)Não foi possível localizar a contra-partida para este evento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
'        End
'    End If
'
'    Geral.EventoLI = RstMDI!Evento
    
    Caixa.Estacao = Val(PegarOpcaoINI("Diversos", "Estacao", "1"))
    Caixa.VersaoAtual = Val(PegarOpcaoINI("Caixa", "Versao", "1"))
    Caixa.CaixaIni = Val(PegarOpcaoINI("Caixa", "Inicial", "1"))
    Caixa.CaixaFim = Val(PegarOpcaoINI("Caixa", "Final", "1"))
           
   'Leitura da data_servidor UBB, tipo_agencia e valor_limite para aprovação
    Set RstUBB = UBBQuery.getControle
   
    Parametros.DataServer = RstUBB("ctlfsdtmvat")                     'DDMMAA sera usada no Calcula_NSU
    Parametros.TipoAgencia = RstUBB("ctlfstpagen")                    'Tipo de Agencia
    
    ValorAlcada = RstUBB("ctlfsvllicx")                               'Valor limite do caixa (Cheques p/aprov. do supervisor-alcada)
    AgenciaUBB = RstUBB("ctlfscdagen")                                'Agencia do ubb
        
    RstUBB.Close
    
    Set RstMDI = MDIQuery.getCargaTabelas(Geral.DataProcessamento)
                
    If RstMDI.EOF Then
        MDIQuery.insParametro Geral.DataProcessamento
    End If
        
    CargaTabelasOK = IIf(RstMDI.EOF, "N", RstMDI("CargaTabelas"))
    
    RstMDI.Close
            
   'Leitura do valor vindo do SQL(MDI image) para transformar ch.saque em ch.compensação
    Set RstMDI = MDIQuery.getParametro(Geral.DataProcessamento)
        
    Parametros.ValorCompensaEnvelope = RstMDI!ValorCompensa_Env
    Parametros.ValorCompensaMaloteVelho = RstMDI!ValorCompensa_Mal
    Parametros.CompensaMaloteNovo = RstMDI!ValorCompensaNovo_Mal
    Parametros.AgenciaCentral = RstMDI!AgenciaCentral
    
    RstMDI.Close
    
   'Verifica se a agencia cadastrada no SQL é a mesma do UBB-NT
    If AgenciaUBB <> Val(Parametros.AgenciaCentral) Then
        MsgBox "Atenção! Verifique o número da agência que está cadastrada no Banco de Dados (tab. agencia central), pois está diferindo do número da agência do Unibanco.", vbOKOnly + vbCritical, "Atenção"
        End
    End If
         
   'leitura do prazo de vencimento das Cobranças Unibanco
    Set RstUBB = UBBQuery.getPrzVenctoVrAlcada(Parametros.AgenciaCentral)
    
    PrzVencimento = RstUBB!agefsnudivc
    Parametros.ValorLimiteInferior = formata(RstUBB!agefsvllisu, True)
    RstUBB.Close
    
   'impressão de dados na tela
    LabelDataServidor = Left(Parametros.DataServer, 2) & "/" & Mid(Parametros.DataServer, 3, 2) & "/20" & Right(Parametros.DataServer, 2)
    LabelMultiAgencia = Parametros.AgenciaCentral
    LabelNoEstacao = Caixa.Estacao
        
    If CargaTabelasOK = "N" Then
        MDIQuery.updParametroAlcada Parametros.ValorLimiteInferior, _
                                                        ValorAlcada, _
                                                        ValorAlcada, _
                                                        PrzVencimento, _
                                                        PrzVencimento, _
                                                        Geral.DataProcessamento
                                                        
       'Atualiza a tabela parametros com o valor lido do dbUbb
        Screen.MousePointer = 0
        If MsgBox("Iniciar a Carga das Tabelas.", vbOKCancel + vbInformation, "Caixa Robô") = vbCancel Then
             End
        End If
                         
       'frmCarga.IniciaCargaTabelas
        frmCarga.Show vbModal
        
        TimerEnter.Enabled = True
        
    Else
       TimerEnter.Enabled = True
    End If
                 
    Screen.MousePointer = 0
    
    Exit Sub
    
TrataErro:
    
    Select Case TratamentoErro("Não foi possível fazer a inicialização do Robô.", Err)
        Case eSair
            End
        Case eRepetir
           Resume
        Case eContinuar
            Resume Next
    End Select
        
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub TimerEnter_Timer()
   TimerEnter.Enabled = False
   Call cmdComunica_Click
End Sub
