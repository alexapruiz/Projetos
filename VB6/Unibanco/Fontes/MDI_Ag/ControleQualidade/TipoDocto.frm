VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DocumentoDesconhecido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Documento"
   ClientHeight    =   3588
   ClientLeft      =   12
   ClientTop       =   480
   ClientWidth     =   5808
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3588
   ScaleWidth      =   5808
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstEscolha 
      Height          =   3468
      Left            =   96
      TabIndex        =   0
      Top             =   48
      Width           =   5628
      _ExtentX        =   9927
      _ExtentY        =   6117
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "Genéricos [F6]"
      TabPicture(0)   =   "TipoDocto.frx":0000
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
      TabPicture(1)   =   "TipoDocto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optTributos(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optTributos(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optTributos(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "optTributos(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optTributos(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optTributos(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optTributos(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Diversos [F8]"
      TabPicture(2)   =   "TipoDocto.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optDiversos(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "optDiversos(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "optDiversos(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "optDiversos(0)"
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
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
Private m_IdEnv_Mal         As String * 1
Private m_Status_Documento  As String * 1
Private Sub cmdFinalizar_Click()
    
    Cancelou = True
    Unload Me
End Sub

Public Function SetIdEnv_Mal(ByVal pIdEnv_Mal As String) As Boolean

    On Error GoTo Erro_IdEnv_Mal:
    
    SetIdEnv_Mal = False
    
    m_IdEnv_Mal = pIdEnv_Mal
    
    SetIdEnv_Mal = True
    
Erro_IdEnv_Mal:

End Function

Public Function SetStatusDocumento(ByVal pStatusDocto As String) As Boolean

    On Error GoTo Erro_StatusDocumento
    
    SetStatusDocumento = False
    
    m_Status_Documento = pStatusDocto
    
    SetStatusDocumento = True
    
Erro_StatusDocumento:
    
End Function


Public Function ShowModal(ByRef pTipoDocto As enumTipoDocto) As Boolean


    Me.Show vbModal
    
    If Not Cancelou Then pTipoDocto = TipoDocto
    
    ShowModal = (Not Cancelou)

End Function

Private Sub Form_Activate()
    Dim i As Byte
    
    Cancelou = False
    Supervisor = False
    
    DocumentoAnterior = False
    
    For i = 0 To optTributos.Count - 1: optTributos(i).Value = False: Next
    For i = 0 To optGenericos.Count - 1: optGenericos(i).Value = False: Next
    For i = 0 To optDiversos.Count - 1: optDiversos(i).Value = False: Next
    
    'Desabilita Opção 'Lançamento Interno' para ENVELOPE
    If m_IdEnv_Mal = "E" Then
        optGenericos(11).Enabled = False
    End If
    'FGTS - Desabilitado
    optTributos(6).Enabled = True
    
    'Força apresentação da pasta de Genéricos
    sstEscolha.SetFocus
    sstEscolha.Tab = 0

    'Desabilita Botão de Docto Anterior eventualmente para complementação
    'do primeiro Documento passado para Capa Envelope/Malote Anterior
'    If Complementacao.grdDocumentos.Row = 0 Then cmdDoctoAnterior.Enabled = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim ikeyDown As Integer
  
  Select Case KeyCode
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
    Unload Me
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
    
    Unload Me
    
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
    Unload Me
End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

