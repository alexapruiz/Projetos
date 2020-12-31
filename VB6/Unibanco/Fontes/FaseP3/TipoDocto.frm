VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form DocumentoDesconhecido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Documento"
   ClientHeight    =   3588
   ClientLeft      =   12
   ClientTop       =   480
   ClientWidth     =   12204
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3588
   ScaleWidth      =   12204
   Begin VB.Frame fraBotoesInferiores 
      Height          =   1356
      Left            =   6864
      TabIndex        =   32
      Top             =   1824
      Width           =   4860
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   800
         Left            =   2016
         Picture         =   "TipoDocto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   336
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
         Left            =   2880
         Picture         =   "TipoDocto.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   336
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
         Left            =   3744
         Picture         =   "TipoDocto.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   336
         Width           =   864
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
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
         Left            =   288
         Picture         =   "TipoDocto.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
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
         Left            =   1152
         Picture         =   "TipoDocto.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
   End
   Begin VB.Frame fraBotoesSuperiores 
      Height          =   1356
      Left            =   7728
      TabIndex        =   31
      Top             =   0
      Width           =   3996
      Begin VB.CommandButton cmdAuditoria 
         Caption         =   "A&uditoria"
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
         Left            =   1152
         Picture         =   "TipoDocto.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "&Finalizar Digitação"
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
         Left            =   2880
         Picture         =   "TipoDocto.frx":10BC
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
      Begin VB.CommandButton cmdSupervisor 
         Caption         =   "Docto &Ilegível"
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
         Left            =   2016
         Picture         =   "TipoDocto.frx":13C6
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
      Begin VB.CommandButton cmdDoctoAnterior 
         Caption         =   "&Docto Anterior"
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
         Left            =   288
         Picture         =   "TipoDocto.frx":16D0
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   336
         Width           =   850
      End
   End
   Begin TabDlg.SSTab sstEscolha 
      Height          =   3468
      Left            =   672
      TabIndex        =   0
      Top             =   48
      Width           =   5628
      _ExtentX        =   9927
      _ExtentY        =   6117
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "Genéricos [F6]"
      TabPicture(0)   =   "TipoDocto.frx":1B12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "optGenericos(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optGenericos(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optGenericos(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "optGenericos(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "optGenericos(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "optGenericos(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "optGenericos(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optGenericos(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optGenericos(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optGenericos(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optGenericos(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optGenericos(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Tributos [F7]"
      TabPicture(1)   =   "TipoDocto.frx":1B2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optTributos(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optTributos(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optTributos(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "optTributos(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optTributos(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optTributos(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optTributos(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Diversos [F8]"
      TabPicture(2)   =   "TipoDocto.frx":1B4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optDiversos(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "optDiversos(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "optDiversos(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "optDiversos(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.OptionButton optGenericos 
         Caption         =   "(C) Lançamento Interno"
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
         Height          =   336
         Index           =   11
         Left            =   600
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2976
         Width           =   2892
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(7) FGTS - Com código de Barras"
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
         Height          =   348
         Index           =   6
         Left            =   -74400
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2496
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(B) CAPA DE OCT"
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
         Height          =   336
         Index           =   10
         Left            =   600
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2736
         Width           =   2892
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(A) OCT"
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
         Height          =   336
         Index           =   9
         Left            =   600
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2496
         Width           =   2892
      End
      Begin VB.OptionButton optDiversos 
         Caption         =   "(4) Arrecadação Convencional"
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
         Height          =   252
         Index           =   3
         Left            =   -74400
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3672
      End
      Begin VB.OptionButton optDiversos 
         Caption         =   "(3) Unicobrança Especial"
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
         Height          =   252
         Index           =   2
         Left            =   -74400
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1776
         Width           =   3672
      End
      Begin VB.OptionButton optDiversos 
         Caption         =   "(2) Unicobrança Registrada"
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
         Height          =   252
         Index           =   1
         Left            =   -74400
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1272
         Width           =   3672
      End
      Begin VB.OptionButton optDiversos 
         Caption         =   "(1) Títulos outros Bcos sem Cod.Barras"
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
         Height          =   252
         Index           =   0
         Left            =   -74400
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   768
         Width           =   3912
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(5) DARM"
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
         Height          =   444
         Index           =   4
         Left            =   -74400
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1788
         Width           =   3672
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(4) GARE"
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
         Height          =   396
         Index           =   3
         Left            =   -74400
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1464
         Width           =   3672
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(2) DARF - Preto"
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
         Height          =   300
         Index           =   1
         Left            =   -74400
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   876
         Width           =   3672
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(1) DARF - Simples"
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
         Height          =   252
         Index           =   0
         Left            =   -74400
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   576
         Width           =   3672
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(3) FGTS"
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
         Height          =   348
         Index           =   2
         Left            =   -74400
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1164
         Width           =   3672
      End
      Begin VB.OptionButton optTributos 
         Caption         =   "(6) GPS"
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
         Height          =   540
         Index           =   5
         Left            =   -74400
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2076
         Width           =   1848
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(7) Autorização de Débito"
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
         Height          =   252
         Index           =   6
         Left            =   600
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1776
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(6) Capa de Envelope"
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
         Height          =   252
         Index           =   5
         Left            =   600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1536
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(5) Código de Barras com Valor de Referência"
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
         Height          =   252
         Index           =   4
         Left            =   600
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1296
         Width           =   4236
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(4) Ficha de Compensação"
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
         Height          =   252
         Index           =   3
         Left            =   600
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1056
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(3) Concessionária"
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
         Height          =   252
         Index           =   2
         Left            =   600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   816
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(2) Depósito (cc - cp)"
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
         Height          =   252
         Index           =   1
         Left            =   600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   576
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(1) Cheque"
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
         Height          =   252
         Index           =   0
         Left            =   600
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   336
         Width           =   3672
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(8) Cartão Crédito Avulso"
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
         Height          =   336
         Index           =   7
         Left            =   600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2892
      End
      Begin VB.OptionButton optGenericos 
         Caption         =   "(9) Capa de Malote"
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
         Height          =   336
         Index           =   8
         Left            =   600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2232
         Width           =   2892
      End
   End
End
Attribute VB_Name = "DocumentoDesconhecido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TipoDocto            As enumTipoDocto
Public Cancelou             As Boolean
Public Supervisor           As Boolean
Public DocumentoAnterior    As Boolean
Private mForm               As Form

Private Sub cmdAuditoria_Click()

    Call Auditoria
    
    sstEscolha.SetFocus

End Sub

Private Sub cmdFinalizar_Click()
    
    Cancelou = True
    Unload Me
End Sub

Private Sub cmdFrenteVerso_Click()

    mForm.cmdFrenteVerso_Click
    sstEscolha.SetFocus
    
End Sub

Private Sub cmdInverteCor_Click()

    mForm.cmdInverteCor_Click
    sstEscolha.SetFocus
    
End Sub

Private Sub cmdRotacao_Click()
    
    mForm.cmdRotacao_Click
    sstEscolha.SetFocus
    
End Sub

Private Sub cmdSupervisor_Click()
    
    
    'Atualiza Log
'    If Not G_GravaLog(Geral.DataProcessamento, Geral.Capa.Capa, Geral.Usuario, "Enviou " & _
'        IIf(Geral.Capa.IdEnv_Mal = "E", "Envelope", "Malote") & " para o Supervisor") Then
'        MsgBox "Não foi possível atualizar arquivo de log!", vbExclamation + vbOKOnly, App.Title
'        Exit Sub
'    End If
    
    frmMotivoDoctoIlegiveis.Show vbModal, Me
    If frmMotivoDoctoIlegiveis.m_CodigoMotivo = 0 Then
        Exit Sub
    End If
    
    If Not GravaMotivoIlegiveisDocto(frmMotivoDoctoIlegiveis.m_CodigoMotivo) Then
        MsgBox "Não foi possível atualizar o motivo do documento para ilegíveis, tente novamente.", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
    Geral.Documento.Status = "5"
    Supervisor = True
    
    Me.Hide

End Sub

Private Sub cmdDoctoAnterior_Click()
    
    Dim iDoctoAtual As Integer, sStatus As String, iTpDocto As Integer

    DocumentoAnterior = False

    'Se Documento anterior é Envelope/Malote, não permite retornar documento
    If Complementacao.grdDocumentos.Row <= 1 Then GoTo SairEnvelope
    
    'Guarda posição atual do documento no grid de navegação
    iDoctoAtual = Complementacao.grdDocumentos.Row
    
    Do While True
        Complementacao.grdDocumentos.Row = (Complementacao.grdDocumentos.Row - 1)

        If Complementacao.grdDocumentos.Row = 0 Then
            MsgBox "Fim de documentos para retorno." + vbCrLf + vbCrLf + "O sistema irá retornar ao próximo documento à complementar.", vbInformation, "Atenção"
            'Primeiro documento em complementação é CAPA, não visualiza e nem permite alteração
            DocumentoAnterior = True: InibirOpcoes (-1)
            GoTo Sair
        End If
        'Obtem status do documento anterior
        Complementacao.grdDocumentos.Col = Complementacao.iColStatus: sStatus = Complementacao.grdDocumentos.Text
        
        'Posiciona somente em documentos (1)Complementados
        If sStatus = "1" Or sStatus = "0" Then Exit Do
    Loop
    
    'Obtem o tipo de documento anterior
    Complementacao.grdDocumentos.Col = Complementacao.iColTpDocto
    
    'Habilita somente a opção do Documento na pasta de Tipos de Documentos
    InibirOpcoes (Complementacao.grdDocumentos.Text)
    
'    cmdSupervisor.Enabled = False
    DocumentoAnterior = True

Sair:
    If Not DocumentoAnterior Then InibirOpcoes (-1) 'Habilita todas opções da pasta
    Me.Hide
    Exit Sub

SairEnvelope:
    'Posiciona no primeiro documento para tratamento no form Digitação
    Complementacao.grdDocumentos.Row = 0
    InibirOpcoes (-1)
    DocumentoAnterior = True
    MsgBox "O sistema só retorna para formulários de documentos." + vbCrLf + vbCrLf + "Não será possível retornar para Capa de Envelope/Malote.", vbInformation, "Atenção"
    GoTo Sair

End Sub

Private Sub cmdZoomMais_Click()
    
    mForm.cmdZoomMais_Click
    sstEscolha.SetFocus
End Sub

Private Sub cmdZoomMenos_Click()
    
    mForm.cmdZoomMenos_Click
    sstEscolha.SetFocus
    
End Sub

Private Sub Form_Activate()
    Dim i As Byte
    
    If Not DocumentoDesconhecido.Visible Then Exit Sub
    
    Cancelou = False
    Supervisor = False
    
    DocumentoAnterior = False
    
    For i = 0 To optTributos.Count - 1: optTributos(i).Value = False: Next
    For i = 0 To optGenericos.Count - 1: optGenericos(i).Value = False: Next
    For i = 0 To optDiversos.Count - 1: optDiversos(i).Value = False: Next
    
    'Desabilita Opção 'Lançamento Interno' para ENVELOPE
    If Geral.Capa.IdEnv_Mal = "E" Then
        optGenericos(11).Enabled = False
'    Else
'        If Geral.capa.IdEnv_Mal = "M" Then
'           optGenericos(11).Enabled = True
'        End If
    End If
    
    'FGTS - Desabilitado
'    optTributos(6).Enabled = True
    
    'Força apresentação da pasta de Genéricos

    sstEscolha.SetFocus
    sstEscolha.Tab = 0

    'Desabilita Botão de Docto Anterior eventualmente para complementação
    'do primeiro Documento passado para Capa Envelope/Malote Anterior
    If Complementacao.grdDocumentos.Row = 0 Then cmdDoctoAnterior.Enabled = False

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim ikeyDown As Integer
  
  Select Case KeyCode
    Case vbKeyAdd
      mForm.cmdZoomMais_Click
    Case vbKeySubtract
      mForm.cmdZoomMenos_Click
    Case vbKeyF10
      mForm.cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      mForm.cmdRotacao_Click
    Case vbKeyF11
      mForm.cmdFrenteVerso_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        mForm.Form_KeyUp KeyCode, Shift
    Case vbKeyF6
        If sstEscolha.TabEnabled(0) = True Then
            sstEscolha.Tab = 0
        End If
    Case vbKeyF7
        If sstEscolha.TabEnabled(1) = True Then
            sstEscolha.Tab = 1
        End If
    Case vbKeyF8
        If sstEscolha.TabEnabled(2) = True Then
            sstEscolha.Tab = 2
        End If
    Case vbKey1 To vbKey9, vbKeyA, vbKeyB, vbKeyC
            If KeyCode = vbKeyA Or KeyCode = vbKeyB Or KeyCode = vbKeyC Then
                ikeyDown = (KeyCode - 56)
            Else
                ikeyDown = KeyCode - 49
            End If

        If sstEscolha.Tab = 0 Then
            If optGenericos(ikeyDown).Enabled Then
                optGenericos_Click (ikeyDown)
                KeyCode = 0
            End If
        ElseIf sstEscolha.Tab = 1 Then
            If KeyCode > vbKey7 Then Exit Sub
            If optTributos(ikeyDown).Enabled Then
                optTributos_Click (ikeyDown)
                KeyCode = 0
            End If
        Else
            If KeyCode > vbKey4 Then Exit Sub
            If optDiversos(ikeyDown).Enabled Then
                optDiversos_Click (ikeyDown)
                KeyCode = 0
            End If
        End If
    Case vbKeyNumpad1 To vbKeyNumpad9
         ikeyDown = (KeyCode - 97)
        
        If sstEscolha.Tab = 0 Then
            If optGenericos(ikeyDown).Enabled Then
                optGenericos_Click (ikeyDown)
                KeyCode = 0
            End If
        ElseIf sstEscolha.Tab = 1 Then
            If KeyCode > vbKeyNumpad7 Then Exit Sub
            If optTributos(ikeyDown).Enabled Then
                optTributos_Click (ikeyDown)
                KeyCode = 0
            End If
        Else
            If KeyCode > vbKeyNumpad4 Then Exit Sub
            If optDiversos(ikeyDown).Enabled Then
                optDiversos_Click (ikeyDown)
                KeyCode = 0
            End If
        End If
    Case vbKeyEscape
            KeyCode = 0
            cmdFinalizar_Click
  End Select

  KeyCode = 0

End Sub

Private Sub optDiversos_Click(Index As Integer)
    Select Case Index
        Case 0  'Títulos de Outros Bancos Sem Cod. Barras
            TipoDocto = etpdocTitulos
        Case 1  'Unicobrança Registrada
            TipoDocto = etpdocCobRegistrada
        Case 2  'Unicobrança Especial
            TipoDocto = etpdocCobEspecial
        Case 3  'Arrecadação Convencional
            TipoDocto = etpdocArrecConvencional
    End Select
    Me.Hide
End Sub

Private Sub optGenericos_Click(Index As Integer)
    Select Case Index
        Case 0  'Cheque
            TipoDocto = etpdocChequeUBBSacado
        Case 1  'Depósito (cc - cp)
            TipoDocto = etpdocDepositoCC
        Case 2  'Concessionária
            TipoDocto = etpdocConcessionariaValorReais
        Case 3  'Ficha de Compensação
            TipoDocto = etpdocFichaCompensacao
        Case 4  'Código de Barras com Vlr Referência
            TipoDocto = etpdocConcessionariaValorIndexado
        Case 5  'Capa de Envelope
            TipoDocto = etpdocEnvelope
        Case 6  'Aviso de Débito
            TipoDocto = etpdocADCC
        Case 7 'Cartâo de Crédito
            TipoDocto = etpdocCartaoAvulso
        Case 8  'Capa de Malote
            TipoDocto = etpdocMalote
        Case 9  '(A) OCT
            TipoDocto = etpdocOCT
        Case 10 '(B) Capa de OCT
            TipoDocto = etpdocCapaOCT
        Case 11 '(C) Lancamento Interno
            TipoDocto = etpdocLancamentoInterno
    End Select
    
    Me.Hide
    
End Sub

Private Sub optTributos_Click(Index As Integer)
    Select Case Index
        Case 0  'Darf - Simples
            TipoDocto = etpdocDarfSimples
        Case 1  'Darf - Preto
            TipoDocto = etpdocDarfPreto
        Case 2  ' FGTS - Arrecadação Convencional
            TipoDocto = etpdocArrecConvencional
        Case 3  'GARE
            TipoDocto = etpdocGare
        Case 4  'DARM
            TipoDocto = etpdocDarm
        Case 5  'GPS
            TipoDocto = etpdocGPS
        Case 6  'FGTS com  Código de Barras
            TipoDocto = etpdocFGTS
    End Select
    Me.Hide
End Sub
Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

Public Sub InibirOpcoes(iTpDocto As Integer)

'Se iTpDocto = (-1) Habilita todos controles das pastas

'Verifica qual documento pertence ao Documento em alteração

'Verifica se docto é uma concessionária
If iTpDocto = etpdocAgua Or iTpDocto = etpdocGas Or _
    iTpDocto = etpdocLuz Or iTpDocto = etpdocTelefone Then
    iTpDocto = etpdocConcessionariaValorReais
End If

'Verifica se docto é uma ficha de compensação
If iTpDocto = etpdocUnicobrancaUBB Or iTpDocto = etpdocCobrancaImediataUBB Or _
    iTpDocto = etpdocCobrancaEspecialUBB Or iTpDocto = etpdocCobrancaTerceiros Then
    iTpDocto = etpdocFichaCompensacao
End If

'Opções da Pasta de Genéricos
If iTpDocto = etpdocChequeUBBSacado Or _
    iTpDocto = etpdocChequeTerceiroPagto Or _
    iTpDocto = etpdocChequeDeposito Or _
    iTpDocto = etpdocDepositoCC Or _
    iTpDocto = etpdocDepositoCP Or _
    iTpDocto = etpdocConcessionariaValorReais Or _
    iTpDocto = etpdocAgua Or _
    iTpDocto = etpdocGas Or _
    iTpDocto = etpdocLuz Or _
    iTpDocto = etpdocTelefone Or _
    iTpDocto = etpdocFichaCompensacao Or _
    iTpDocto = etpdocConcessionariaValorIndexado Or _
    iTpDocto = etpdocTributosMunicipais Or _
    iTpDocto = etpdocTributosEstaduais Or _
    iTpDocto = etpdocTributosFederais Or _
    iTpDocto = etpdocEnvelope Or _
    iTpDocto = etpdocADCC Or _
    iTpDocto = etpdocMalote Or _
    iTpDocto = etpdocOCT Or _
    iTpDocto = etpdocCapaOCT Or _
    iTpDocto = etpdocCartaoAvulso Or _
    iTpDocto = etpdocLancamentoInterno Or _
    iTpDocto = -1 Or _
    iTpDocto = -9 Then

    If iTpDocto = -1 Then
        sstEscolha.TabEnabled(0) = True
    Else
        sstEscolha.TabEnabled(0) = True: sstEscolha.TabEnabled(1) = False: sstEscolha.TabEnabled(2) = False
        SendKeys "{F6}"
    End If

    'Cheque UBB, Terceiros ou Depósito
    optGenericos(0).Enabled = (iTpDocto = etpdocChequeUBBSacado Or _
                               iTpDocto = etpdocChequeTerceiroPagto Or _
                               iTpDocto = etpdocChequeDeposito Or iTpDocto = -1)
    'Depósito (cc - cp)
    optGenericos(1).Enabled = (iTpDocto = etpdocDepositoCC Or _
                               iTpDocto = etpdocDepositoCP Or iTpDocto = -1)
    'Concessionária
    optGenericos(2).Enabled = (iTpDocto = etpdocConcessionariaValorReais Or _
                               iTpDocto = etpdocAgua Or iTpDocto = etpdocGas Or _
                               iTpDocto = etpdocLuz Or iTpDocto = etpdocTelefone Or iTpDocto = -1)
    'Ficha de Compensação
    optGenericos(3).Enabled = (iTpDocto = etpdocFichaCompensacao Or iTpDocto = -1)
    'Código Barras com Valor Indexado
    optGenericos(4).Enabled = (iTpDocto = etpdocConcessionariaValorIndexado Or _
                               iTpDocto = etpdocTributosEstaduais Or _
                               iTpDocto = etpdocTributosFederais Or _
                               iTpDocto = etpdocTributosMunicipais Or iTpDocto = -1)
    'Capa de Envelope
    optGenericos(5).Enabled = (iTpDocto = etpdocEnvelope Or iTpDocto = -1 Or iTpDocto = -9)
    'ADCC Autorização de Débito
    optGenericos(6).Enabled = (iTpDocto = etpdocADCC Or iTpDocto = -1)
    'Cartâo de Crédito
    optGenericos(7).Enabled = (iTpDocto = etpdocCartaoAvulso Or iTpDocto = -1)
    'Capa de Malote
    optGenericos(8).Enabled = (iTpDocto = etpdocMalote Or iTpDocto = -1 Or iTpDocto = -9)
    'OCT
    optGenericos(9).Enabled = (iTpDocto = etpdocOCT Or iTpDocto = -1)
    'Capa de Oct
    optGenericos(10).Enabled = (iTpDocto = etpdocCapaOCT Or iTpDocto = -1)
    
    'Lançamento Interno (Habilita opção somente para MALOTE)
    If Geral.Capa.IdEnv_Mal = "E" Then
        optGenericos(11).Enabled = False
    Else
        optGenericos(11).Enabled = (iTpDocto = etpdocLancamentoInterno Or iTpDocto = -1)
    End If
    
End If

'Opções da Pasta de Tributos
If iTpDocto = etpdocDarfSimples Or _
    iTpDocto = etpdocDarfPreto Or _
    iTpDocto = etpdocArrecConvencional Or _
    iTpDocto = etpdocFGTS Or _
    iTpDocto = etpdocGare Or _
    iTpDocto = etpdocDarm Or _
    iTpDocto = etpdocGPS Or _
    iTpDocto = -1 Then
    
    If iTpDocto = -1 Then
        sstEscolha.TabEnabled(1) = True
    Else
        sstEscolha.TabEnabled(0) = False: sstEscolha.TabEnabled(1) = True: sstEscolha.TabEnabled(2) = False
        sstEscolha.Tab = 1
        SendKeys "{F7}"
    End If
    optTributos(0).Enabled = (iTpDocto = etpdocDarfSimples Or iTpDocto = -1)        'Darf - Simples
    optTributos(1).Enabled = (iTpDocto = etpdocDarfPreto Or iTpDocto = -1)          'Darf - Preto
    optTributos(2).Enabled = (iTpDocto = etpdocArrecConvencional Or iTpDocto = -1)  'FGTS (SEM Cód Barras - Arrecadação Convencional
    optTributos(3).Enabled = (iTpDocto = etpdocGare Or iTpDocto = -1)               'GARE
    optTributos(4).Enabled = (iTpDocto = etpdocDarm Or iTpDocto = -1)               'DARM
    optTributos(5).Enabled = (iTpDocto = etpdocGPS Or iTpDocto = -1)                'GPS
    optTributos(6).Enabled = (iTpDocto = etpdocFGTS Or iTpDocto = -1)               'FGTS (Com Código de Barras)
    
End If

'Opções da Pasta de Diversos
If iTpDocto = etpdocTitulos Or _
    iTpDocto = etpdocCobRegistrada Or _
    iTpDocto = etpdocCobEspecial Or _
    iTpDocto = etpdocArrecConvencional Or _
    iTpDocto = -1 Then
    
    If iTpDocto = -1 Then
        sstEscolha.TabEnabled(2) = True
    Else
        sstEscolha.TabEnabled(0) = False: sstEscolha.TabEnabled(1) = False: sstEscolha.TabEnabled(2) = True
        SendKeys "{F8}"
    End If
    
    optDiversos(0).Enabled = (iTpDocto = etpdocTitulos Or iTpDocto = -1)            'Títulos de Outros Bancos Sem Cod. Barras
    optDiversos(1).Enabled = (iTpDocto = etpdocCobRegistrada Or iTpDocto = -1)      'Unicobrança Registrada
    optDiversos(2).Enabled = (iTpDocto = etpdocCobEspecial Or iTpDocto = -1)        'Unicobrança Especial
    optDiversos(3).Enabled = (iTpDocto = etpdocArrecConvencional Or iTpDocto = -1)  'Arrecadação Convencional
    
End If

'If iTpDocto = -1 Then cmdSupervisor.Enabled = True
Me.cmdDoctoAnterior.Enabled = (iTpDocto <> -9)

End Sub

Private Sub sstEscolha_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        mForm.Form_KeyUp KeyCode, Shift
    End If
    
End Sub
Private Function GravaMotivoIlegiveisDocto(ByVal lCodMotivo As Long) As Boolean

Dim qryMotivoIlegiveis  As rdoQuery
    
GravaMotivoIlegiveisDocto = False

On Error GoTo Err_GravaMotivoIlegiveisDocto
    
    Screen.MousePointer = vbHourglass
    
    Set qryMotivoIlegiveis = Geral.Banco.CreateQuery("", "{? = call  AlteraMotivoIlegiveisDocto(?,?,?)}")
    With qryMotivoIlegiveis
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto
        .rdoParameters(3) = lCodMotivo
        .Execute
    
        If .rdoParameters(0).Value <> 0 Then GoTo Exit_GravaMotivoIlegiveisDocto
    End With
    
    GravaMotivoIlegiveisDocto = True

Exit_GravaMotivoIlegiveisDocto:
    
    Screen.MousePointer = vbDefault
    If Not (qryMotivoIlegiveis Is Nothing) Then qryMotivoIlegiveis.Close
    
    Exit Function

Err_GravaMotivoIlegiveisDocto:
    
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do código de motivos para documento ilegível." & vbCrLf & Err.Description, Err, rdoErrors)
        Case vbCancel, vbRetry
    End Select
    GoTo Exit_GravaMotivoIlegiveisDocto

End Function
