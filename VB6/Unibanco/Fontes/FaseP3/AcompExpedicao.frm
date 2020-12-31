VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form AcompExpedicao 
   AutoRedraw      =   -1  'True
   Caption         =   "Acompanhamento de Expedição"
   ClientHeight    =   6504
   ClientLeft      =   360
   ClientTop       =   1512
   ClientWidth     =   11772
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6504
   ScaleWidth      =   11772
   Begin VB.Frame Frame3 
      Height          =   732
      Left            =   5652
      TabIndex        =   45
      Top             =   5688
      Width           =   6024
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Confirmar"
         Height          =   384
         Left            =   648
         TabIndex        =   0
         Top             =   252
         Width           =   1512
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   384
         Left            =   3756
         TabIndex        =   2
         Top             =   252
         Width           =   1512
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   384
         Left            =   2196
         TabIndex        =   1
         Top             =   252
         Width           =   1512
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Escolha uma opção de Gráfico"
      Height          =   636
      Left            =   5832
      TabIndex        =   37
      Top             =   288
      Width           =   5688
      Begin VB.OptionButton Expedido 
         Caption         =   "Expedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3204
         TabIndex        =   39
         Top             =   288
         Width           =   1452
      End
      Begin VB.OptionButton optNexpedido 
         Caption         =   "A Expedir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1188
         TabIndex        =   38
         Top             =   288
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      Height          =   528
      Left            =   240
      TabIndex        =   8
      Top             =   5688
      Width           =   5184
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   828
         TabIndex        =   11
         Top             =   216
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Envelopes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1968
         TabIndex        =   10
         Top             =   204
         Width           =   1236
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Malotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3396
         TabIndex        =   9
         Top             =   204
         Width           =   1236
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   300
      Left            =   1080
      ScaleHeight     =   252
      ScaleWidth      =   2064
      TabIndex        =   6
      Top             =   156
      Width           =   2112
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Data do Movimento"
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
         Height          =   192
         Left            =   24
         TabIndex        =   7
         Top             =   24
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   300
      Left            =   3204
      ScaleHeight     =   252
      ScaleWidth      =   1212
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   156
      Width           =   1260
      Begin VB.Label lblDataProc 
         Alignment       =   2  'Center
         Caption         =   "17/07/2000"
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
         Height          =   204
         Left            =   48
         TabIndex        =   5
         Top             =   24
         Width           =   1152
      End
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   5340
      Left            =   144
      OleObjectBlob   =   "AcompExpedicao.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   252
      Width           =   5196
   End
   Begin VB.Label lblTotCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8868
      TabIndex        =   44
      Top             =   4284
      Width           =   732
   End
   Begin VB.Label lblTotPorCapa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9648
      TabIndex        =   43
      Top             =   4284
      Width           =   492
   End
   Begin VB.Label lblTotDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   10176
      TabIndex        =   42
      Top             =   4284
      Width           =   876
   End
   Begin VB.Label lblTotPorDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   11076
      TabIndex        =   41
      Top             =   4284
      Width           =   492
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   6228
      TabIndex        =   40
      Top             =   4284
      Width           =   2604
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6384
      Left            =   108
      TabIndex        =   12
      Top             =   36
      Width           =   5448
   End
   Begin VB.Label lblPorCapaProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9624
      TabIndex        =   35
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label lblQtdDocProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   10164
      TabIndex        =   34
      Top             =   2280
      Width           =   876
   End
   Begin VB.Label lblPorDocProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   11064
      TabIndex        =   33
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label lblQtdCapaProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8856
      TabIndex        =   32
      Top             =   2280
      Width           =   732
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   6216
      TabIndex        =   31
      Top             =   2268
      Width           =   2604
   End
   Begin VB.Label lblCor 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   0
      Left            =   5784
      TabIndex        =   30
      Top             =   2280
      Width           =   408
   End
   Begin VB.Label lblQtdCapaPProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8856
      TabIndex        =   29
      Top             =   2940
      Width           =   732
   End
   Begin VB.Label lblPorDocPProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   11064
      TabIndex        =   28
      Top             =   2940
      Width           =   492
   End
   Begin VB.Label lblQtdDocPProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   10164
      TabIndex        =   27
      Top             =   2940
      Width           =   876
   End
   Begin VB.Label lblPorCapaPProc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9624
      TabIndex        =   26
      Top             =   2940
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parcialmente Processado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   6216
      TabIndex        =   25
      Top             =   2940
      Width           =   2604
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   1
      Left            =   5760
      TabIndex        =   24
      Top             =   2940
      Width           =   408
   End
   Begin VB.Label lblQtdCapaRej 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8856
      TabIndex        =   23
      Top             =   3612
      Width           =   732
   End
   Begin VB.Label lblPorDocRej 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   11064
      TabIndex        =   22
      Top             =   3612
      Width           =   492
   End
   Begin VB.Label lblQtdDocRej 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   10164
      TabIndex        =   21
      Top             =   3612
      Width           =   876
   End
   Begin VB.Label lblPorCapaRej 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9624
      TabIndex        =   20
      Top             =   3612
      Width           =   492
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rejeitado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   6216
      TabIndex        =   19
      Top             =   3612
      Width           =   2604
   End
   Begin VB.Label lblCor 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Index           =   2
      Left            =   5784
      TabIndex        =   18
      Top             =   3612
      Width           =   408
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   6216
      TabIndex        =   17
      Top             =   1596
      Width           =   2604
   End
   Begin VB.Label lblFiltro 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Capas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8856
      TabIndex        =   16
      Top             =   1596
      Width           =   732
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9624
      TabIndex        =   15
      Top             =   1596
      Width           =   492
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Doctos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   10164
      TabIndex        =   14
      Top             =   1596
      Width           =   876
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   11064
      TabIndex        =   13
      Top             =   1596
      Width           =   492
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   5628
      Left            =   5652
      TabIndex        =   36
      Top             =   36
      Width           =   6036
   End
End
Attribute VB_Name = "AcompExpedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* Type de Não Expedidos
    Private Type NExpedidos
        TotalCapa           As Long
        TotalDocs           As Long
        Rejeitado           As Long
        Processado          As Long
        PProcessado         As Long
        DocRejeitado        As Long
        DocProcessado       As Long
        DocPProcessado      As Long
    End Type
    Private NExpedidos As NExpedidos
    
'* Type de Expedidos
    Private Type Expedidos
        TotalCapa           As Long
        TotalDocs           As Long
        Rejeitado           As Long
        Processado          As Long
        PProcessado         As Long
        DocRejeitado        As Long
        DocProcessado       As Long
        DocPProcessado      As Long
    End Type
    Private Expedidos As Expedidos
    
Private qryAcompExpedido     As rdoQuery     '-- query de Expedidos (Malote/Envelope)
Private qryAcompExpedidoE    As rdoQuery     '-- query de Expedidos (Envelope)
Private qryAcompExpedidoM    As rdoQuery     '-- query de Expedidos (Malote)

Private RsExpedidos          As rdoResultset '-- recordsets

Private Status               As String       '-- Status - 'T' ou 'E'
Private Index                As Integer      '-- quantidade de partes do Gráfico
Private IdEnvMal             As String       '-- identificador de Envelope/Malote/Todos
Private Graficos             As Boolean
Private Sub CmdFechar_Click()
'* Sai da Tela de Acompanhamento de Expedicao *'
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
'* Imprime tela Atual
    Screen.MousePointer = vbHourglass
    Me.PrintForm
    Screen.MousePointer = vbDefault
End Sub
Private Sub CmdOK_Click()
'-- Faz query para trazer dados --'
On Error GoTo TrataErro

    Screen.MousePointer = vbHourglass

    If IdEnvMal = "T" Then
    
        With qryAcompExpedido
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = Status
            Set RsExpedidos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
    ElseIf IdEnvMal = "E" Then
    
        With qryAcompExpedidoE
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = Status
            Set RsExpedidos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
    Else
    
        With qryAcompExpedidoM
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = Status
            Set RsExpedidos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
    End If
    
    If Not RsExpedidos.EOF Then
        'Não Expedidos
        If Status = "T" Then
            NExpedidos.Processado = RsExpedidos!Processado
            NExpedidos.DocProcessado = RsExpedidos!DocProcessado
            
            RsExpedidos.MoreResults
            
            NExpedidos.PProcessado = RsExpedidos!ParcProcessado
            NExpedidos.DocPProcessado = RsExpedidos!DocParcProcessado
            
            RsExpedidos.MoreResults
            
            NExpedidos.Rejeitado = RsExpedidos!Rejeitado
            NExpedidos.DocRejeitado = RsExpedidos!DocRejeitado
            
            '--Fim de Tratamento caso todos os campo sejam = 0--3'
            If NExpedidos.Processado = 0 And NExpedidos.DocProcessado = 0 _
                And NExpedidos.PProcessado = 0 And NExpedidos.DocPProcessado = 0 _
                And NExpedidos.Rejeitado = 0 And NExpedidos.DocRejeitado = 0 Then
                Call LimpaLabels
                Call LimpaGrafico
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
           'Expedidos
            Expedidos.Processado = RsExpedidos!Processado
            Expedidos.DocProcessado = RsExpedidos!DocProcessado
            
            RsExpedidos.MoreResults
            
            Expedidos.PProcessado = RsExpedidos!ParcProcessado
            Expedidos.DocPProcessado = RsExpedidos!DocParcProcessado
            
            RsExpedidos.MoreResults
            
            Expedidos.Rejeitado = RsExpedidos!Rejeitado
            Expedidos.DocRejeitado = RsExpedidos!DocRejeitado

            '--Fim de Tratamento caso todos os campo sejam = 0--3'
            If Expedidos.Processado = 0 And Expedidos.DocProcessado = 0 _
                And Expedidos.PProcessado = 0 And Expedidos.DocPProcessado = 0 _
                And Expedidos.Rejeitado = 0 And Expedidos.DocRejeitado = 0 Then
                Screen.MousePointer = vbDefault
                Call LimpaLabels
                Call LimpaGrafico
                Exit Sub
            End If
        End If
    End If
            
    Screen.MousePointer = vbDefault
    
    '* Fecha Recordset *'
    RsExpedidos.Close

    Call PreencheGrafico
    Call PreencheLabels
    
Exit Sub
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível atualizar Gráfico.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub
Private Sub Expedido_Click()
Call LimpaLabels
Call LimpaGrafico

'-- Marca Status = 'E' - Expedido --'
    Status = "E"
       
End Sub
Private Sub Form_Activate()
   
   '* Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(27)
    
   '* Data de Processamento Atual
   lblDataProc.Caption = Mid(Geral.DataProcessamento, 7, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 5, 2) & "/" & _
                         Mid(Geral.DataProcessamento, 1, 4)
    
End Sub
Private Sub Form_Load()

    '--query de Expedidos / Não Expedidos (Envelope/Malote)
    Set qryAcompExpedido = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicao(?,?)}")

    '--query de Expedidos / Não Expedidos (Envelope)
    Set qryAcompExpedidoE = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicaoE(?,?)}")

    '--query de Expedidos / Não Expedidos (Malote)
    Set qryAcompExpedidoM = Geral.Banco.CreateQuery("", "{Call GetAcompExpedicaoM(?,?)}")
    
    '--default
    Call LimpaGrafico
    
    '--status default 'T'
    Status = "T"
    
    '--Identificação default 'Todos'
    IdEnvMal = "T"
    
End Sub
Private Sub LimpaGrafico()
'--Definição de formato default do Gráfico--'
On Error GoTo TrataErro

    With Grafico
        .Column = 1
        .Data = Index
        .Column = 2
        .Data = Index
        .Column = 3
        .Data = Index
    End With
    
    Graficos = False

Exit Sub
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível Limpar Gráfico.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    
End Sub
Private Sub lblCor_Click(Index As Integer)
'-- Mostra numero de capa para cada opção --'
On Error GoTo TrataErro
    
    If Graficos = False Then Exit Sub
        ListaCapa.m_idtela = 1
        ListaCapa.m_tpGrafico = Index
        ListaCapa.m_InStatus = Status
        ListaCapa.m_IdEnv_Mal = IdEnvMal
        ListaCapa.Show vbModal, Me
        
Exit Sub
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível Listar Capas.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
        
End Sub
Private Sub optFiltro_Click(Index As Integer)
Call LimpaLabels
Call LimpaGrafico
'-- 0 Todos / 1 Envelope / 2 Malote --'
    If Index = 0 Then
        IdEnvMal = "T"
    ElseIf Index = 1 Then
        IdEnvMal = "E"
    Else
        IdEnvMal = "M"
    End If
End Sub
Private Sub PreencheGrafico()
'--Preenche Gráfico com informaçoes obtidas na query--'
On Error GoTo TrataErro

    Graficos = True

    If Status = "T" Then
        '--Não Expedido
        With Grafico
            .Column = 1
            .Data = NExpedidos.Processado
            .Column = 2
            .Data = NExpedidos.PProcessado
            .Column = 3
            .Data = NExpedidos.Rejeitado
        End With
    
    Else
        '--Expedido
        With Grafico
            .Column = 1
            .Data = Expedidos.Processado
            .Column = 2
            .Data = Expedidos.PProcessado
            .Column = 3
            .Data = Expedidos.Rejeitado
        End With
    
    End If
    
Exit Sub
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível Preencher Gráfico.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
        
End Sub
Private Sub PreencheLabels()
'--Preenche Labels com informaçoes obtidas na query--'
On Error GoTo TrataErro

    If Status = "T" Then
        '--Não Expedido (Capa)
        lblQtdCapaProc.Caption = Format(NExpedidos.Processado, "#,##0")
        lblQtdCapaPProc.Caption = Format(NExpedidos.PProcessado, "#,##0")
        lblQtdCapaRej.Caption = Format(NExpedidos.Rejeitado, "#,##0")
        NExpedidos.TotalCapa = (NExpedidos.Processado + NExpedidos.PProcessado + NExpedidos.Rejeitado)
        lblTotCapa.Caption = Format(NExpedidos.TotalCapa, "#,#00")
        
        '--Não Expedido (Documentos)
        lblQtdDocProc.Caption = Format(NExpedidos.DocProcessado, "#,#00")
        lblQtdDocPProc.Caption = Format(NExpedidos.DocPProcessado, "#,#00")
        lblQtdDocRej.Caption = Format(NExpedidos.DocRejeitado, "#,#00")
        NExpedidos.TotalDocs = (NExpedidos.DocProcessado + NExpedidos.DocPProcessado + NExpedidos.DocRejeitado)
        lblTotDoc.Caption = Format(NExpedidos.TotalDocs, "#,#00")
        
        '--Não Expedido Porcentagem(Capa)
        lblPorCapaProc = Format(NExpedidos.Processado * 100 / NExpedidos.TotalCapa, "0.0")
        lblPorCapaPProc = Format(NExpedidos.PProcessado * 100 / NExpedidos.TotalCapa, "0.0")
        lblPorCapaRej = Format(NExpedidos.Rejeitado * 100 / NExpedidos.TotalCapa, "0.0")
        
        '--Não Expedido Porcentagem(Documento)
        lblPorDocProc = Format(NExpedidos.DocProcessado * 100 / NExpedidos.TotalDocs, "0.0")
        lblPorDocPProc = Format(NExpedidos.DocPProcessado * 100 / NExpedidos.TotalDocs, "0.0")
        lblPorDocRej = Format(NExpedidos.DocRejeitado * 100 / NExpedidos.TotalDocs, "0.0")
    Else
        '--Expedido (Capa)
        lblQtdCapaProc.Caption = Format(Expedidos.Processado, "#,#00")
        lblQtdCapaPProc.Caption = Format(Expedidos.PProcessado, "#,#00")
        lblQtdCapaRej.Caption = Format(Expedidos.Rejeitado, "#,#00")
        Expedidos.TotalCapa = (Expedidos.Processado + Expedidos.PProcessado + Expedidos.Rejeitado)
        lblTotCapa.Caption = Format(Expedidos.TotalCapa, "#,#00")
        
        '--Expedido (Documentos)
        lblQtdDocProc.Caption = Format(Expedidos.DocProcessado, "#,#00")
        lblQtdDocPProc.Caption = Format(Expedidos.DocPProcessado, "#,#00")
        lblQtdDocRej.Caption = Format(Expedidos.DocRejeitado, "#,#00")
        Expedidos.TotalDocs = (Expedidos.DocProcessado + Expedidos.DocPProcessado + Expedidos.DocRejeitado)
        lblTotDoc.Caption = Format(Expedidos.TotalDocs, "#,#00")
        
        '--Expedido Porcentagem(Capa)
        lblPorCapaProc = Format(Expedidos.Processado * 100 / Expedidos.TotalCapa, "0.0")
        lblPorCapaPProc = Format(Expedidos.PProcessado * 100 / Expedidos.TotalCapa, "0.0")
        lblPorCapaRej = Format(Expedidos.Rejeitado * 100 / Expedidos.TotalCapa, "0.0")
        
        '--Expedido Porcentagem(Documento)
        lblPorDocProc = Format(Expedidos.DocProcessado * 100 / Expedidos.TotalDocs, "0.0")
        lblPorDocPProc = Format(Expedidos.DocPProcessado * 100 / Expedidos.TotalDocs, "0.0")
        lblPorDocRej = Format(Expedidos.DocRejeitado * 100 / Expedidos.TotalDocs, "0.0")
    
    End If

        '--Não Expedido Total Porcentagem(Capa)
        lblTotPorCapa = CalcPorcentagem(lblPorCapaProc, lblPorCapaPProc, lblPorCapaRej)
        
        '--Não Expedido Total Porcentagem(Documento)
        lblTotPorDoc = CalcPorcentagem(lblPorDocProc, lblPorDocPProc, lblPorDocRej)

Exit Sub
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível atualizar Labels.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
            
End Sub
Function CalcPorcentagem(string1 As String, string2 As String, string3 As String) As String
'* Calcula a Soma de 3 Porcentagens *'
On Error GoTo TrataErro

Dim avstring1   As Integer 'string 1 antes  da vírgula
Dim avstring2   As Integer 'string 2 antes  da vírgula
Dim avstring3   As Integer 'string 3 antes  da vírgula

Dim dvstring1   As Integer 'string 1 depois da vírgula
Dim dvstring2   As Integer 'string 2 depois da vírgula
Dim dvstring3   As Integer 'string 3 depois da vírgula

Dim totalav     As Integer 'total da string antes  da vírgula
Dim totaldv     As Integer 'total da string depois da vírgula

    avstring1 = Mid(string1, 1, InStr(1, string1, ",", 1) - 1)
    avstring2 = Mid(string2, 1, InStr(1, string2, ",", 1) - 1)
    avstring3 = Mid(string3, 1, InStr(1, string3, ",", 1) - 1)

    dvstring1 = Mid(string1, InStr(1, string1, ",", 1) + 1, Len(string1))
    dvstring2 = Mid(string2, InStr(1, string2, ",", 1) + 1, Len(string2))
    dvstring3 = Mid(string3, InStr(1, string3, ",", 1) + 1, Len(string3))

    totalav = (avstring1 + avstring2 + avstring3)
    totaldv = (dvstring1 + dvstring2 + dvstring3)
        
    If totalav = 99 And totaldv = 10 Then
        CalcPorcentagem = "100,0"
    ElseIf totalav = 99 And totaldv = 100 Then
        CalcPorcentagem = "100,0"
    Else
        CalcPorcentagem = (totalav & "," & totaldv)
    End If
            
Exit Function
TrataErro:
Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível atualizar Gráfico.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
            
End Function
Private Sub optNexpedido_Click()
Call LimpaLabels
Call LimpaGrafico
'-- Marca Status = 'T' - Não Expedido --'
    Status = "T"
End Sub
Private Sub LimpaLabels()
'--Limpa Labels--'
    lblQtdCapaProc.Caption = "0"
    lblQtdCapaPProc.Caption = "0"
    lblQtdCapaRej.Caption = "0"
    lblTotCapa.Caption = "0"
    
    lblQtdDocProc.Caption = "0"
    lblQtdDocPProc.Caption = "0"
    lblQtdDocRej.Caption = "0"
    lblTotDoc.Caption = "0"
    
    lblPorCapaProc.Caption = "0.0"
    lblPorCapaPProc.Caption = "0.0"
    lblPorCapaRej.Caption = "0.0"
    
    lblPorDocProc.Caption = "0.0"
    lblPorDocPProc.Caption = "0.0"
    lblPorDocRej.Caption = "0.0"

    lblTotPorCapa.Caption = "0.0"
    lblTotPorDoc.Caption = "0.0"
End Sub
